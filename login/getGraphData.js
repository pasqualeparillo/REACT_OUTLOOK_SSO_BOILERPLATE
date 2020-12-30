/* global OfficeRuntime */

import axios from "axios";
import { getGraphToken } from "./loginHelper";
let retryGetAccessToken = 0;

export async function getGraphDataTest() {
  try {
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true
    });
    let exchangeResponse = await getGraphToken("/api", bootstrapToken);
    console.log("this is bootgrap token");
    console.log(exchangeResponse);
    console.log("this is exchange response token");
  } catch (exception) {
    console.log("bootstrap token failed");
    console.log(exception);
  }
}
