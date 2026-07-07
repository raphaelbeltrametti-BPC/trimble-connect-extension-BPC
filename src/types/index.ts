export interface TCModell {
  id: string;
  name: string;
}

export interface TCProjekt {
  id: string;
  name: string;
  location?: string;
}

export interface ViewerState {
  selektion: number[];
  aktivesModellId: string;
  modelle: TCModell[];
}
