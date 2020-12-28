import React, { useEffect, useState } from "react";
import { signInO365 } from "../../../login/loginHelper";
import { logoutFromO365 } from "../../../login/logoutHelper";
export default function Test() {
  const login = async () => {
    await signInO365();
  };
  const logout = async () => {
    await logoutFromO365();
  };
  return (
    <div>
      <button onClick={() => login()}>LogIn</button>
      <button onClick={() => logout()}>LogOut</button>
    </div>
  );
}
