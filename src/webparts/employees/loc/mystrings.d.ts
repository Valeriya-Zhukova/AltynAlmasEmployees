declare interface IEmployeesWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'EmployeesWebPartStrings' {
  const strings: IEmployeesWebPartStrings;
  export = strings;
}
