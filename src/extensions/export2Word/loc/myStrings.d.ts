declare interface IExport2WordCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'Export2WordCommandSetStrings' {
  const strings: IExport2WordCommandSetStrings;
  export = strings;
}
