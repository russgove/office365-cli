import {textToTextFieldTransformer} from "./textToTextFieldTransformer";
import {IFieldTransformer} from './IFieldTransformer';
export interface ITransformerDefinition{
  fromFieldType:string,
   toFieldType:string,
   name:string,
   transformer: IFieldTransformer,
   description:string
}

var transfomers:Array<ITransformerDefinition>=[
  {fromFieldType:"Text", toFieldType:"Text",transformer:new textToTextFieldTransformer(),name:"TextToText",description:"Text to Text-- can be used to change the internal name of a field"}
]
export default transfomers;