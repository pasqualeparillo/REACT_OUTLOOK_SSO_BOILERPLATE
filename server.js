const express = require("express");
let app = express();

app.get("/", (req, res) => {
  const authorization = req.get("Authorization");
  if (authorization == null) {
    let error = new Error("No Authorization header was found.");
    res.send(error);
  } else {
    res.send(authorization);
  }
});
app.get("/api", (req, res) => res.send("HELLO FROM EXPRESS API"));
app.get("/auth", (req, res) => res.send("HELLO FROM EXPRESS AUTH"));
app.use(express.static("public"));
app.listen(5000, () => console.log("Example app listening on port 5000!"));
