declare interface IRbStampUiCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'RbStampUiCommandSetStrings' {
  const strings: IRbStampUiCommandSetStrings;
  export = strings;
}
