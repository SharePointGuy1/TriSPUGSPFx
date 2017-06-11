import { ISPList } from './HelloTriSpugWebPart';

export default class MockHttpClient {

    private static _items: ISPList[] = [
        { Title: 'Mock List', Id: '1', BaseType: 0 },
        { Title: 'Mock List 2', Id: '2', BaseType: 1 },
        { Title: 'Mock List 3', Id: '3', BaseType: 0 }
    ];

    public static get(): Promise<ISPList[]> {
        return new Promise<ISPList[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }
}