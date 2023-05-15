import {
  ListViewCommandSetContext
} from '@microsoft/sp-listview-extensibility';

// import pnp and pnp logging system
import { spfi, SPFI, SPFx as spSPFx } from "@pnp/sp";
import { GraphFI, graphfi, SPFx as graphSPFx } from "@pnp/graph";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp"
import "@pnp/sp/folders";
import "@pnp/sp/webs";
import "@pnp/sp/sites";
import "@pnp/sp/webs/types";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/graph/users";
import "@pnp/graph/teams";
import "@pnp/graph/sites/group";
import "@pnp/graph/groups";

var _sp: SPFI = null;
var _graph: GraphFI = null;

export const getSP = (context?: ListViewCommandSetContext): SPFI => {
  if (_sp === null && context != null) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _sp = spfi().using(spSPFx(context)).using(PnPLogging(LogLevel.Warning));
  }
  return _sp;
};

export const getGraph = (context?: ListViewCommandSetContext): GraphFI => {
  if (_graph === null && context != null) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _graph = graphfi().using(graphSPFx(context)).using(PnPLogging(LogLevel.Warning));
  }
  return _graph;
};