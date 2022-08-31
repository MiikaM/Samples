import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseAdaptiveCardExtension } from "@microsoft/sp-adaptive-card-extension-base";
import { CardView } from "./cardView/CardView";
import { QuickView } from "./quickView/QuickView";
import { HtmlFeedCardPropertyPane } from "./HtmlFeedCardPropertyPane";
import { HtmlReader } from "../models/HtmlReader";
import HtmlFeedReaderService from "../Service/HtmlFeedReaderService";

export interface IHtmlFeedCardAdaptiveCardExtensionProps {
  title: string;
}

export interface IHtmlFeedCardAdaptiveCardExtensionState {}

const CARD_VIEW_REGISTRY_ID: string = "HtmlFeedCard_CARD_VIEW";
export const QUICK_VIEW_REGISTRY_ID: string = "HtmlFeedCard_QUICK_VIEW";

export default class HtmlFeedCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IHtmlFeedCardAdaptiveCardExtensionProps,
  IHtmlFeedCardAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: HtmlFeedCardPropertyPane | undefined;

  public onInit(): Promise<void> {
    console.log("Init");

    let _htmlContent: HtmlReader;
    const htmlReaderService: HtmlFeedReaderService = new HtmlFeedReaderService();
    return htmlReaderService.getHtmlContent(_htmlContent, this.context.httpClient).then((outputContent: any) => {
      console.log({outputContent});

      // this.state = {
      //   ID: 1,
      //   SearchText: "",
      //   Items: ouputPodcastsContent.channel.item,
      //   channel: ouputPodcastsContent.channel,
      // };
      this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
      this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());
      return Promise.resolve();

    });
    // this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    // this.quickViewNavigator.register(
    //   QUICK_VIEW_REGISTRY_ID,
    //   () => new QuickView()
    // );

    // return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'HtmlFeedCard-property-pane'*/
      "./HtmlFeedCardPropertyPane"
    ).then((component) => {
      this._deferredPropertyPane = new component.HtmlFeedCardPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }
}
