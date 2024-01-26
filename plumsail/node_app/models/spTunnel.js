import { AuthConfig } from "node-sp-auth-config";
import { bootstrap } from "pnp-auth";
import pkg from '@pnp/sp-commonjs';
const { sp, Web } = pkg;


function spAuthentication(webURL){
  let authConfig = new AuthConfig({ configPath: "./config/creds_ntlm.json", encryptPassword: true, saveConfigOnDisk: true });
  bootstrap(sp, authConfig, webURL);
}

export async function getItem(webURL, listname, columns, filter){
  spAuthentication(webURL);
   var _item = sp.web.lists
              .getByTitle(listname)
              .items
             .select(columns);

             if(filter != '')
             _item.filter(filter);

    return await _item.get()
                .then((items) => {
                  return items;
                });
}

export async function addItem(webURL, listname, data){
  spAuthentication(webURL);
  return await sp.web.lists
              .getByTitle(listname)
              .items
              .add(data);
}



