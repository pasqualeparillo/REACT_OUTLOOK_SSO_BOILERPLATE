import React, { useEffect, useState } from "react";
import { getGraphDataTest } from "../../../login/getGraphData";
import { signInO365 } from "../../../login/loginHelper";
export default function Test() {
  const [token, setToken] = useState("");
  const [error, setError] = useState("");
  const login = async () => {
    await signInO365(displayError, setState);
  };

  const displayError = e => {
    setError.set(e);
  };
  async function setState(t) {
    await setToken(t);
  }

  return (
    <div>
      <button onClick={() => login()}>LogIn</button>
      <button onClick={() => getGraphDataTest()}>GetGraphData</button>
      <p>{token.toString()}</p>
    </div>
  );
}
