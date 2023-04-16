/// <reference path="../../node_modules/@types/office-js/index.d.ts" />
import Common from "./common";
/* global window Office */
export default class Signatures {
  #error = null;
  #user_info = null;
  /**
   * Loads the users signature information
   * @param {SignatureUserInfo} [user_info]
   */
  constructor(user_info = null) {
    if (!user_info) {
      /** @type {String} */
      let user_info_str = window.localStorage.getItem("user_info");
      if (user_info_str) {
        this.#user_info = JSON.parse(user_info_str);
      } else {
        // -- display notification alert here --
        this.#error = "Unable to retrieve 'user_info' from localStorage.";
      }
    } else {
      this.#user_info = user_info;
    }
  }
  /**
   * Returns the last error
   * @returns {String}
   */
  get last_error() {
    return this.#error;
  }
  /**
   * Load template A
   * @param {String} fn
   * @returns {String}
   */
  get_template_A = (fn = null) => {
    let str = "";
    if (Common.is_valid_data(this.#user_info.greeting)) {
      str += this.#user_info.greeting + "<br/>";
    }
    let imgInsert = `<img src="https://officedev.github.io/Office-Add-in-samples/Samples/outlook-set-signature/assets/sample-logo.png" alt="Logo" />`;
    if (fn === null) {
      imgInsert = `<img src="cid:${fn}" alt="MS Logo" width="24" height="24" />`;
    }
    str += `<table>
              <tr>
                <td style='border-right: 1px solid #000000; padding-right: 5px;'>
                  ${imgInsert}
                </td>
                <td style='padding-left: 5px;'>
                  <b>${this.#user_info.name.trim()}</b>
                  ${Common.is_valid_data(this.#user_info.pronoun) ? "&nbsp;" + this.#user_info.pronoun : ""}<br/>
                  ${Common.is_valid_data(this.#user_info.job) ? this.#user_info.job + "<br/>" : ""}
                  ${this.#user_info.email}<br/>
                  ${Common.is_valid_data(this.#user_info.phone) ? this.#user_info.phone + "<br/>" : ""}
                </td>
              </tr>
            </table>`;
    return str;
  };
  /**
   * Loads template B
   * @returns {String}
   */
  get_template_B = () => {
    let str = "";
    if (Common.is_valid_data(this.#user_info.greeting)) {
      str += this.#user_info.greeting + "<br/>";
    }
    str += `<table>
              <tr>
                <td style='border-right: 1px solid #000000; padding-right: 5px;'>
                  <img src='https://officedev.github.io/Office-Add-in-samples/Samples/outlook-set-signature/assets/sample-logo.png' alt='Logo' />
                </td>
                <td style='padding-left: 5px;'>
                  <strong>${this.#user_info.name}</strong>
                  ${Common.is_valid_data(this.#user_info.pronoun) ? "&nbsp;" + this.#user_info.pronoun : ""}<br/>
                  ${this.#user_info.email}<br/>
                  ${Common.is_valid_data(this.#user_info.phone) ? this.#user_info.phone + "<br/>" : ""}
                </td>
              </tr>
          </table>`;
    return str;
  };
  /**
   * Returns template C
   * @returns {String}
   */
  get_template_C = () => {
    let str = "";
    if (Common.is_valid_data(this.#user_info.greeting)) {
      str += this.#user_info.greeting + "<br/>";
    }
    str += this.#user_info.name;
    return str;
  };
  /**
   * Returns the settings from the user
   * @returns {SignatureSettings}
   */
  get_signature_settings = () => {
    /** @type {SignatureSettings} */
    const returnValue = {};
    let val = Office.context.roamingSettings.get("newMail");
    if (val) returnValue.newMail = val;
    val = Office.context.roamingSettings.get("reply");
    if (val) returnValue.reply = val;
    val = Office.context.roamingSettings.get("forward");
    if (val) returnValue.forward = val;
    val = Office.context.roamingSettings.get("override_olk_signature");
    if (val != null) returnValue.override = val;
    return returnValue;
  };
}
/**
 * @typedef {Object} SignatureSettings
 * @property {String} newMail
 * @property {String} reply
 * @property {String} forward
 * @property {Boolean} override
 */
/**
 * @typedef {Object} SignatureUserInfo
 * @property {String} greeting
 * @property {String} name
 * @property {String} email
 * @property {String} phone
 * @property {String} job
 * @property {String} pronoun
 */
