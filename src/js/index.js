/// <reference path="../../node_modules/@types/office-js/index.d.ts" />
/* global Office  document require  */
const { default: SignatureSettingsInterface } = require("./signatureSettingsInterface");
// eslint-disable-next-line office-addins/no-office-initialize
Office.initialize = () => {
  // load the stylesheet
  const linkElement = document.createElement("link");
  linkElement.rel = "stylesheet";
  linkElement.type = "text/css";
  linkElement.href = "default.css";
  document.head.appendChild(linkElement);
  // load the form
  const container = document.getElementById("container");
  const settingsForm = new SignatureSettingsInterface();
  settingsForm.render(container);
};
