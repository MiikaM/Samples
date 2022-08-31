// import axios from "axios";
import axios from 'axios';
import cheerio from "cheerio";
import { HtmlReader } from "../models/HtmlReader";

const HTML_FEED_URL = "https://<tenant>.sharepoint.com/Shared%20Documents/<page>/feed.html";

export default class HtmlFeedReaderService  {

    public getHtmlContent = async(_htmlContent: HtmlReader, httpClient: any): Promise<any> => {
        try {
            // const response: any = await httpClient.get(
            //     // `https://feed.podbean.com/pnpweekly/feed.xml`,
            //     HTML_FEED_URL,
            //     HttpClient.configurations.v1);

            const response: any = await axios.get(HTML_FEED_URL);
            
            const data: any = response?.data;

            const dom = cheerio.load(data);
            const feed = [];
            const hitList = dom('#hit-list').children('table').first();
            hitList.each((idx, el) => {
                const taa = dom(el); 
                const feedInfo = {title: '', url: '', image: null};
                feedInfo.title = dom(el).children('tbody').children('tr').first().children('td').children('h3').text();
                feedInfo.url = dom(el).children('tbody').children('tr').first().children('td').children('h3').children('a').attr('href');
                var image = taa.find('table').first().find('image').attr('src');

            })
            // console.log({htmlcontent});
            
            return data;
        } catch(error) {
            console.error({error});
        }
    }
}
