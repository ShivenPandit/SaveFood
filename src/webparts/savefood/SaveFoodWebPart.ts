import * as React from "react";
import * as ReactDom from "react-dom/client";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import "../../tailwind.generated.css";
import * as strings from "TemplateSpfxWebPartStrings";
import SaveFood from "./components/SaveFood";
import { ISaveFoodProps } from "./components/ISaveFoodProps";
import { DEFAULT_PORTAL_URL } from "./config/portalConfig";
import $ from "jquery";

export interface ISaveFoodWebPartProps {
  description: string;
}

interface ISharePointUserGroup {
  Id: number;
  Title: string;
}

interface IUserData {
  id: number;
  title: string;
  email: string;
  loginName: string;
  isSiteAdmin: boolean;
  groups: ISharePointUserGroup[];
}

export default class SaveFoodWebPart extends BaseClientSideWebPart<ISaveFoodWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _root: ReactDom.Root | null = null;
  private _userData: IUserData = {
    id: -1,
    title: "",
    email: "",
    loginName: "",
    isSiteAdmin: false,
    groups: [],
  };
  private _domObserver: MutationObserver | null = null;
  private _injectedStyleElement: HTMLStyleElement | null = null;
  private _isApplyingDomFixes: boolean = false;

  public async getCurrentUserGroups(
    requester: SPHttpClient,
    siteUrl: string,
  ): Promise<IUserData> {
    try {
      const response: SPHttpClientResponse = await requester.get(
        `${siteUrl}/_api/web/currentuser/?$select=Title,Email,LoginName,Id,IsSiteAdmin&$expand=groups`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "odata-version": "",
          },
        },
      );

      if (!response.ok) {
        throw new Error(`HTTP ${response.status} - ${response.statusText}`);
      }

      const json = await response.json();
      const user = json?.d;

      return {
        id: user?.Id ?? -1,
        title: user?.Title ?? "",
        email: this.context.pageContext.user.email || user?.Email || "",
        loginName: user?.LoginName || this.context.pageContext.user.loginName || "",
        isSiteAdmin: !!user?.IsSiteAdmin || !!(this.context.pageContext.legacyPageContext as any)?.isSiteAdmin,
        groups: user?.Groups?.results ?? [],
      };
    } catch (error: unknown) {
      console.error(
        "[SPFx] Error fetching current user groups:",
        error instanceof Error ? error.message : error,
      );
      return {
        id: -1,
        title: "",
        email: this.context.pageContext.user.email || "",
        loginName: this.context.pageContext.user.loginName || "",
        isSiteAdmin: !!(this.context.pageContext.legacyPageContext as any)?.isSiteAdmin,
        groups: [],
      };
    }
  }

  protected async onInit(): Promise<void> {
    await super.onInit();

    // 1. Synchronously inject CSS rules and DOM observers before any async operations
    this._injectHideStyles();
    this._setupDomObserver();

    // Re-enforce full-bleed layout whenever window is resized
    window.addEventListener("resize", () => {
      this._enforceFullBleed();
    });

    document.title = "Save Food Management System";
    $("body").addClass("ms-backgroundImage");

    // 2. Fetch user data before rendering component
    this._userData = await this.getCurrentUserGroups(
      this.context.spHttpClient,
      this.context.pageContext.web.absoluteUrl,
    );
  }

  public render(): void {
    this.domElement.style.setProperty("display", "flex", "important");
    this.domElement.style.setProperty("flex-direction", "column", "important");
    this.domElement.style.setProperty("opacity", "1", "important");
    this.domElement.style.setProperty("visibility", "visible", "important");
    this.domElement.style.setProperty("width", "100%", "important");
    this.domElement.style.setProperty("max-width", "100%", "important");
    this.domElement.style.setProperty("min-width", "100%", "important");
    this.domElement.style.setProperty("min-height", "100vh", "important");
    this.domElement.style.setProperty("box-sizing", "border-box", "important");
    this.domElement.style.setProperty("margin", "0px", "important");
    this.domElement.style.setProperty("padding", "0px", "important");
    this.domElement.style.setProperty("background-color", "#F8FAFC", "important");

    this._enforceFullBleed();

    const element: React.ReactElement<ISaveFoodProps> =
      React.createElement(SaveFood, {
        newcontext: this.context,
        newsiteUrl: this.context.pageContext.web.absoluteUrl || DEFAULT_PORTAL_URL,
        newlistName: "",
        newspHttpClient: this.context.spHttpClient,
        newrelativeUrl: this.context.pageContext.web.serverRelativeUrl,
        newcurrentUserId: this._userData.id,
        newcurrentUserName: this._userData.title,
        newgroups: this._userData.groups,
        newcurrentUserEmail: this._userData.email,
        newcurrentUserLoginName: this._userData.loginName,
        newisSiteAdmin: this._userData.isSiteAdmin,
      });

    // React 18 createRoot
    if (!this._root) {
      this._root = ReactDom.createRoot(this.domElement);
    }
    this._root.render(element);
  }

  private _enforceFullBleed(): void {
    try {
      $(".sp-pageLayout, .sp-pageLayout-sideNav, #spPageChromeAppDiv, #spPageCanvasContent").show();

      let p = this.domElement ? this.domElement.parentElement : null;
      while (p && p !== document.body && p !== document.documentElement) {
        p.style.setProperty("display", "block", "important");
        p.style.setProperty("opacity", "1", "important");
        p.style.setProperty("visibility", "visible", "important");
        p.style.setProperty("width", "100%", "important");
        p.style.setProperty("max-width", "100%", "important");
        p.style.setProperty("min-width", "100%", "important");
        p.style.setProperty("min-height", "100%", "important");
        p.style.setProperty("height", "100%", "important");
        p.style.setProperty("margin", "0px", "important");
        p.style.setProperty("padding", "0px", "important");
        p.style.setProperty("border", "none", "important");
        p.style.setProperty("box-sizing", "border-box", "important");
        p.style.setProperty("background-color", "#F8FAFC", "important");
        p = p.parentElement;
      }

      const canvasElements = document.querySelectorAll(
        '.CanvasZone, .CanvasSection, .ControlZone, #spPageCanvasContent, [data-automation-id="CanvasControl"], [data-automation-id="CanvasZone"], [data-automation-id="CanvasSection"], [class*="CanvasSection"], [class*="CanvasZone"], [class*="ControlZone"], [class*="CanvasComponent"]'
      );
      canvasElements.forEach((el) => {
        const htmlEl = el as HTMLElement;
        htmlEl.style.setProperty("display", "block", "important");
        htmlEl.style.setProperty("opacity", "1", "important");
        htmlEl.style.setProperty("visibility", "visible", "important");
        htmlEl.style.setProperty("max-width", "100%", "important");
        htmlEl.style.setProperty("width", "100%", "important");
        htmlEl.style.setProperty("min-width", "100%", "important");
        htmlEl.style.setProperty("min-height", "100%", "important");
        htmlEl.style.setProperty("height", "100%", "important");
        htmlEl.style.setProperty("margin", "0px", "important");
        htmlEl.style.setProperty("padding", "0px", "important");
        htmlEl.style.setProperty("border", "none", "important");
        htmlEl.style.setProperty("box-sizing", "border-box", "important");
        htmlEl.style.setProperty("background-color", "#F8FAFC", "important");
      });
    } catch {
      // ignore
    }
  }

  private _injectHideStyles(): void {
    const selectors = [
      "#spTopPlaceholder",
      ".spTopPlaceholder",
      'div[data-automation-id="spTopPlaceholder"]',
      "#spTopContentPlaceholder",
      'div[data-sp-placeholder="TopContent"]',
      "#spSiteHeader",
      ".spSiteHeader",
      'div[data-sp-feature-tag="Site header host"]',
      'div[data-automation-id="SiteHeader"]',
      'div[data-automation-id="pageHeader"]',
      'div[data-automation-id="pageHeaderBackground"]',
      "#spCommandBar",
      'div[id*="CommandBar"]',
      'div[class*="commandBarWrapper"]',
      ".commandBarWrapper",
      "#spLeftNav",
      "#SuiteNavWrapper",
      "#suiteBarDelta",
      "#suiteBarTop",
      "#suiteBarWrapper",
      "#suiteBar",
      "#O365_NavHeader",
      "#mru_nav_header",
      "#headerRegion",
      ".o365cs-base",
      ".o365cs-nav-header",
      "div[class*='headerRegion']",
      "div[class*='suiteNav']",
      "div[class*='suiteBar']",
      "div[id*='suiteBar']",
      "div[id*='SuiteNav']",
      "div[id*='O365']",
      "div[class*='o365']",
      "[data-automation-id='searchBox']",
      "#O365_SearchBoxContainer",
      "#searchBoxPositioner",
      "#CommentsWrapper",
      "#sp-appBar",
      "#spAppBArNGSPFlyoutContainer",
      "[data-automation-id='sp-appBar']",
      ".headerRow-107",
      ".commandBarButtonHeightAndColor",
      ".fui-Toolbar",
      ".simpleFooterContainer-196",
      ".p_i-Dej_1x34n",
      'div[data-automation-id="MegaFooter"]',
      'div[data-automation-id="SimpleFooter"]',
      'div[data-automation-id="pageBottom"]',
      'div[data-automation-id="footer"]',
      'div[data-automation-id="PageFooter"]',
      'footer[role="contentinfo"]',
      '.sp-page-footer',
      'div[data-automation-id="gradientBox"]',
      'div[data-automation-id="LeftNavGroups"]',
      'div[data-automation-id="SiteHeaderOverlay"]',
      'div[data-automation-id="titleRegionMarker"]',
      "#workbenchCommandBar",
    ];

    this._injectedStyleElement = document.createElement("style");
    this._injectedStyleElement.innerHTML = `
      ${selectors.join(", ")} {
        display: none !important;
        height: 0 !important;
        min-height: 0 !important;
        max-height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
        border: none !important;
        visibility: hidden !important;
        pointer-events: none !important;
        position: absolute !important;
        top: -9999px !important;
      }
      .fixed, [class*="fixed"] {
        margin: 0 !important;
        margin-top: 0 !important;
      }
      html, body {
        overflow-x: hidden !important;
        max-width: 100% !important;
        background-color: #F8FAFC !important;
        font-family: 'Inter', 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif !important;
        -webkit-font-smoothing: antialiased !important;
        -moz-osx-font-smoothing: grayscale !important;
        margin: 0 !important;
        padding: 0 !important;
        margin-top: 0 !important;
        padding-top: 0 !important;
        border: none !important;
        -webkit-overflow-scrolling: touch !important;
      }
      div[data-automation-id="contentScrollRegion"],
      #workbenchPageContent,
      #spPageChromeAppDiv,
      .sp-appRoot,
      .sp-pageLayout,
      .sp-pageLayout-sideNav {
        overflow-y: auto !important;
        overflow-x: hidden !important;
        height: 100% !important;
        min-height: 100% !important;
        width: 100% !important;
        max-width: 100% !important;
        background-color: #F8FAFC !important;
        margin: 0 !important;
        padding: 0 !important;
        margin-top: 0 !important;
        padding-top: 0 !important;
        top: 0 !important;
        position: static !important;
        transform: none !important;
        border: none !important;
        -webkit-overflow-scrolling: touch !important;
      }
      /* Allow natural vertical scrolling on all screens (tablet, small laptop, desktop) when content exceeds viewport */
      body,
      html,
      body.savefood-dashboard-no-scroll,
      body.savefood-dashboard-no-scroll div[data-automation-id="contentScrollRegion"],
      body.savefood-dashboard-no-scroll #workbenchPageContent,
      body.savefood-dashboard-no-scroll #spPageChromeAppDiv,
      body.savefood-dashboard-no-scroll .sp-appRoot,
      body.savefood-dashboard-no-scroll .sp-pageLayout,
      body.savefood-dashboard-no-scroll #spPageCanvasContent,
      body.savefood-dashboard-no-scroll .savefood-canvas-breakout,
      html.savefood-dashboard-no-scroll,
      div[data-automation-id="contentScrollRegion"],
      #workbenchPageContent,
      #spPageChromeAppDiv,
      .sp-appRoot,
      .sp-pageLayout,
      .sp-pageLayout-sideNav,
      #spPageCanvasContent,
      .savefood-canvas-breakout {
        overflow-y: auto !important;
        overflow-x: hidden !important;
        -webkit-overflow-scrolling: touch !important;
        min-height: 100vh !important;
      }
      #spPageCanvasContent,
      #spPageCanvasContent > div,
      .SPCanvas,
      .SPCanvas-canvas,
      .CanvasZone,
      .CanvasZoneSectionContainer,
      .CanvasSection,
      .ControlZone,
      div[data-automation-id="CanvasControl"],
      div[data-automation-id="CanvasZone"],
      div[data-automation-id="CanvasSection"],
      .CanvasComponent,
      .Canvas,
      [class*="CanvasSection"],
      [class*="CanvasZone"],
      [class*="ControlZone"],
      [class*="CanvasControl"],
      [class*="sp-pageLayout"],
      [class*="pageContent"],
      [class*="canvasContent"],
      [data-sp-feature-tag="SaveFoodWebPart"],
      .ms-Grid,
      .appPart,
      .webPartContainer,
      .root-98,
      .mainContent,
      section.mainContent,
      .b_3Ynvt_1x34n,
      div[class*="b_3Ynvt"],
      .c_TBFAg_St4Iq,
      div[class*="c_TBFAg"],
      div[class*="p_U54gn"],
      div[class*="p_pANWZ"],
      div[class*="p_PjYhq"],
      div[class*="f_g4Tmq"],
      div[class*="w_42xsm"] {
        display: block !important;
        opacity: 1 !important;
        visibility: visible !important;
        max-width: 100% !important;
        width: 100% !important;
        min-width: 100% !important;
        min-height: 100% !important;
        height: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
        border: none !important;
        border-radius: 0 !important;
        box-sizing: border-box !important;
        background-color: #F8FAFC !important;
      }
      /* Fluent UI Layer, Callout & Picker transparency fix */
      .ms-Layer,
      .ms-Layer--fixed,
      .ms-Layer-content,
      .ms-Layer .ms-Fabric,
      .ms-Callout-container,
      .ms-Callout {
        background-color: transparent !important;
        background: transparent !important;
      }
      .ms-Callout-main {
        background-color: #ffffff !important;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1) !important;
        border-radius: 12px !important;
        border: 1px solid #e2e8f0 !important;
      }
      .CanvasZone,
      .CanvasSection,
      .CanvasZoneSectionContainer,
      .ControlZone,
      div[data-automation-id="CanvasControl"],
      div[data-automation-id="CanvasZone"],
      div[data-automation-id="CanvasSection"] {
        padding-bottom: 0 !important;
        margin-bottom: 0 !important;
      }
      .CanvasSection-xl12, .CanvasSection-l12, .CanvasSection-m12, .CanvasSection-sm12, .CanvasSection-col {
        max-width: 100% !important;
        width: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
      }
      .od-SiteNav-content, [class*="Canvas"] {
        margin: 0 !important;
        padding: 0 !important;
        border: none !important;
        max-width: 100% !important;
        width: 100% !important;
        border-radius: 0 !important;
        background-color: #F8FAFC !important;
      }
      .savefood-canvas-breakout {
        width: 100% !important;
        max-width: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
        box-sizing: border-box !important;
        overflow-x: hidden !important;
      }
    `;
    document.head.appendChild(this._injectedStyleElement);
  }

  private _setupDomObserver(): void {
    const hideExtraNodes = (): void => {
      if (this._isApplyingDomFixes) return;
      this._isApplyingDomFixes = true;

      try {
        $(
          "#SuiteNavWrapper, #suiteBarTop, #suiteBarDelta, #suiteBarWrapper, #suiteBar, #O365_NavHeader, #mru_nav_header, #headerRegion, .o365cs-base, .o365cs-nav-header, #spSiteHeader, #spTopPlaceholder, .spTopPlaceholder, #spTopContentPlaceholder, div[data-automation-id='SiteHeader'], div[data-automation-id='pageHeader'], div[data-automation-id='SimpleFooter'], div[data-automation-id='SiteHeaderOverlay'], div[data-automation-id='gradientBox'], div[data-automation-id='LeftNavGroups'], #spCommandBar, .commandBarWrapper, div[id*='CommandBar'], div[class*='commandBarWrapper'], #sp-appBar, [data-automation-id='sp-appBar'], [data-automation-id='searchBox'], #O365_SearchBoxContainer, #searchBoxPositioner",
        ).hide();
      } finally {
        this._isApplyingDomFixes = false;
      }
    };

    hideExtraNodes();

    let debounceTimer: number | null = null;
    this._domObserver = new MutationObserver((mutations) => {
      if (this._isApplyingDomFixes) return;

      let hasAddedNodes = false;
      for (let i = 0; i < mutations.length; i++) {
        // If mutation is occurring inside our SaveFood webpart, ignore it to prevent page-switching lag
        if (this.domElement && this.domElement.contains(mutations[i].target as Node)) {
          continue;
        }
        if (mutations[i].addedNodes && mutations[i].addedNodes.length > 0) {
          hasAddedNodes = true;
          break;
        }
      }
      if (!hasAddedNodes) return;

      if (debounceTimer) {
        window.clearTimeout(debounceTimer);
      }
      debounceTimer = window.setTimeout(() => {
        hideExtraNodes();
      }, 120);
    });

    this._domObserver.observe(document.body, {
      childList: true,
      subtree: true,
    });
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null,
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null,
      );
    }
  }

  protected onDispose(): void {
    if (this._domObserver) {
      this._domObserver.disconnect();
      this._domObserver = null;
    }

    if (this._injectedStyleElement && this._injectedStyleElement.parentNode) {
      this._injectedStyleElement.parentNode.removeChild(
        this._injectedStyleElement,
      );
      this._injectedStyleElement = null;
    }

    if (this._root) {
      this._root.unmount();
      this._root = null;
    }
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
