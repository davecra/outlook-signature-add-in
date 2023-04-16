/// <reference path="../../node_modules/@types/office-js/index.d.ts" />
/* global Office require */
const { default: Common } = require("./common");
const { default: SignatureEventWrapper } = require("./signatureEventWrapper");
// eslint-disable-next-line office-addins/no-office-initialize
Office.initialize = () => {};
// mount the command into Office
Office.actions.associate("checkSignature", () => {
  let user_info_str = Office.context.roamingSettings.get("user_info");
  if (!user_info_str) {
    Common.display_insight_infobar();
  } else {
    let user_info = JSON.parse(user_info_str);
    const wrapper = new SignatureEventWrapper(user_info);
    wrapper.addSignature("default");
  }
});
