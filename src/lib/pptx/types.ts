
export interface TextShape {
    id: string;
    name: string;
    text: string;
    paragraphs: string[];
}

export interface TableContent {
    id: string;
    name: string;
    rows: string[][];
}

export interface SlideContent {
    index: number;
    slideId: string;
    shapes: TextShape[];
    tables: TableContent[];
}

export interface PresentationContent {
    fileName: string;
    slides: SlideContent[];
}

export interface ChangeAuthor {
    name?: string;
    email?: string;
    userId?: string;
}

export interface ChangeMetadata {
    id?: string;
    actionId?: string;
    timestamp?: string;
    author: ChangeAuthor;
}

export interface PresentationChangeRecord {
    operation: string;
    level: "document" | "slide" | "shape" | "graphicFrame";
    metadata: ChangeMetadata;
    slideId?: string;
    shapeId?: string;
}

