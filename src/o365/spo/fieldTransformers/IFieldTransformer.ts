export interface IFieldTransformer {

  /**
   * @param fieldInternalName the internal name of the field to be transformed.
   * 
   * @returns  an array of fields to be added to the selects clause of the request and an array of expands to be 
   * to the $expands clause of the request
   */
  setQuery(fieldInternalName: string): { selects: Array<string>; expands: Array<string> };

  /**
     * @param listitem : The listitem selected from the sharepoint list that will include the fields requested in the 
   * selects and expands returned from setqury.
   * 
   * @param fieldInternalName :The internalName of the field in the listitem used to create the result
   * 
   * @returns an object that can be used to update the target field type
   */
  setJSON(listitem: any, fieldInternalName: string): any;
}