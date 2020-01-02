import request from '../../../request';
import { IFieldTransformer, IFieldDefinition, ITransformerDefinition } from "./fieldTransformers";
export class textToUserFieldTransformer implements IFieldTransformer {
  setQuery(fromFieldDef: IFieldDefinition, transformerDefinition: ITransformerDefinition): { selects: Array<string>; expands: Array<string> } {
    var selects: Array<string> = [fromFieldDef.InternalName];
    var expands: Array<string> = [];
    return { selects: selects, expands: expands }
  }
  async setJSON(listitem: any, fromFieldDef: IFieldDefinition, toFieldDef: IFieldDefinition, transformerDefinition: ITransformerDefinition, webUrl: string, formDigestValue: string): Promise<any> {

    let update: any = {};
    if (transformerDefinition.searchDisplayName) {
      update[`${toFieldDef.InternalName}Id`] = await this.getNumericUserIdUsingName(webUrl, formDigestValue, listitem[fromFieldDef.InternalName]);
    } else {
      update[`${toFieldDef.InternalName}Id`] = await this.getNumericUserId(webUrl, formDigestValue, listitem[fromFieldDef.InternalName]);
    }

    return update;
  }
  private async getNumericUserIdUsingName(webUrl: string, formDigestValue: string, userName: string): Promise<number | null> {
    //maybe add sine caching here one day.....
    const userPrincipalName: string | null = await this.findUserByName(userName, formDigestValue);

    if (userPrincipalName !== null) {
      return this.getNumericUserId(webUrl, formDigestValue, userPrincipalName);
    }
    else { return null }
  }
  private async findUserByName(userDisplayName: string, formDigestValue: string): Promise<string | null> {
    //maybe add sine caching here one day.....
    const findUserByNameQuery: any = {
      url: `https://graph.microsoft.com/v1.0/users?$select=userPrincipalName,displayName&$filter=displayName eq '${userDisplayName}'`,
      headers: {
        'X-RequestDigest': formDigestValue,
        'Content-Type': `application/json;odata=verbose`,
        'Accept': `application/json`,
      }
    }

    const upn: string | null = await request.get(findUserByNameQuery)
      .then((userresult: any) => {
        userresult = JSON.parse(userresult);
        if (userresult.value.length === 0) {
          console.log(`user ${userDisplayName} was not found`);
          return null;
        }
        else {
          if (userresult.value.length > 1) {
            console.log(`${userresult.value.length} users with Display name ${userDisplayName} were found`);
            return null;
          }
        }

        return userresult.value[0].userPrincipalName;
      })
      .catch((err) => {
        console.log(err);

        console.log(`an arror occurred fetching user ${userDisplayName}`);
        return null;
      });

    return upn;

  }
  private async getNumericUserId(webUrl: string, formDigestValue: string, userEmail: string): Promise<number | null> {
    //maybe add sine caching here one day.....
    var logonName = `i:0#.f|membership|${userEmail}`;
    const ensureUserOption: any = {
      url: `${webUrl}/_api/web/ensureuser`,
      headers: {
        'X-RequestDigest': formDigestValue,
        'Content-Type': `application/json;odata=verbose`,
        'Accept': `application/json`,
      },
      body: JSON.stringify({ 'logonName': logonName })
    }
    var id: number | null = null;
    await request.post(ensureUserOption)
      .then((userresult: any) => {
        userresult = JSON.parse(userresult);
        id = userresult.Id;
      })
      .catch((err) => {
        console.log(`user ${userEmail} was not found`);
        id = null;
      });
    return id;

  };

}
