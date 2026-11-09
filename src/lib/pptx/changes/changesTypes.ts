export interface ChangeAuthor {
    name?: string;
    email?: string;
    userId?: string;
    providerId?: string;
}

export interface ChangeMetadata {
    id?: string;
    actionId?: string;
    timestamp?: string;
    deviceId?: string;
    version?: string;
    author: ChangeAuthor;
}

export interface PresentationChangeRecord {
    operation: string;
    level: "document" | "slide" | "shape" | "graphicFrame";
    metadata: ChangeMetadata;
    slideId?: string;
    shapeId?: string;
    graphicFrameId?: string;
}
