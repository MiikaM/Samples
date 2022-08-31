export interface HtmlReader {
    channel: Channel;
}

export interface Channel {
    title: string;
    link: string;
    // image: Image;
    item?: (ItemEntity)[] | null;
}

export interface Image {
    url: string;
}
export interface ItemEntity {
    title: string;
    link: string;
    pubDate: string;
}