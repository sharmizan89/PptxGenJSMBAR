export interface PlaceholderIdxRegistry {
    [layoutName: string]: {
        pic?: number[];
        body?: number[];
        title?: number[];
        ftr?: number[];
        sldNum?: number[];
        dt?: number[];
    };
}
export interface PlaceholderNameRegistry {
    [layoutName: string]: {
        [placeholderName: string]: {
            type: string;
            idx?: number;
        };
    };
}
export declare const PLACEHOLDER_NAME_REGISTRY: PlaceholderNameRegistry;
export declare const PLACEHOLDER_IDX_REGISTRY: PlaceholderIdxRegistry;
