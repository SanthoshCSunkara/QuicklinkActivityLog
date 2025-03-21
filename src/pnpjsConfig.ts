import { SPFI, spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { WebPartContext } from "@microsoft/sp-webpart-base";

let _sp: SPFI;

export const getSP = (context?: WebPartContext): SPFI => {
    if (!_sp && context) {
        _sp = spfi().using(SPFx(context));
    }
    return _sp;
};
