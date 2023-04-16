/// <reference path="../../node_modules/@types/office-js/index.d.ts" />
import Common from "./common";
import Signatures from "./signatures";
/* global Office */
export default class SignatureEventWrapper {
  /** @type {SignatureTemplate} */
  #signature_info = {};
  /** @type {import("./signatures").SignatureUserInfo} */
  #user_info = {};
  /**
   * For Outlook on Windows and on Mac only. Insert signature into appointment or message.
   * Outlook on Windows and on Mac can use setSignatureAsync method on appointments and messages.
   * @param {*} user_info Information details about the user
   * @param {*} eventObj Office event object
   */
  constructor(user_info) {
    this.#user_info = user_info;
  }
  /**
   * Add templates to the signature
   * @param {"templateA" | "templateB" | "templateC" | "default" } which
   */
  addSignature = (which) => {
    if (Office.context.mailbox.item.getComposeTypeAsync) {
      //Find out if the compose type is "newEmail", "reply", or "forward" so that we can apply the correct template.
      Office.context.mailbox.item.getComposeTypeAsync((result) => {
        const templateName = this.#get_template_name(result.value);
        let template_name = which === "default" || which === undefined ? templateName : which;
        const templateHtml = this.#get_signature_info(template_name);
        this.#signature_info = templateHtml;
        this.#insert_user_signature();
      });
    } else {
      if (Office.context.mailbox.item.itemType == "appointment") {
        this.#insertAppointment();
      } else {
        this.#insert_user_signature();
      }
    }
  };
  /**
   * Inserts to an appointment body
   */
  #insertAppointment = () => {
    if (Common.is_valid_data(this.#signature_info.logoBase64) === true) {
      //If a base64 image was passed we need to attach it.
      Office.context.mailbox.item.addFileAttachmentFromBase64Async(
        this.#signature_info.logoBase64,
        this.#signature_info.logoFileName,
        {
          isInline: true,
        },
        () => {
          Office.context.mailbox.item.body.setAsync(Common.get_cal_offset() + this.#signature_info.signature, {
            coercionType: "html",
          });
        }
      );
    } else {
      Office.context.mailbox.item.body.setAsync(Common.get_cal_offset() + this.#signature_info.signature, {
        coercionType: "html",
      });
    }
  };
  /**
   * Inserts the users formatted signature
   */
  #insert_user_signature = () => {
    if (Common.is_valid_data(this.#signature_info.logoBase64) === true) {
      //If a base64 image was passed we need to attach it.
      Office.context.mailbox.item.addFileAttachmentFromBase64Async(
        this.#signature_info.logoBase64,
        this.#signature_info.logoFileName,
        {
          isInline: true,
        },
        () => {
          //After image is attached, insert the signature
          Office.context.mailbox.item.body.setSignatureAsync(this.#signature_info.signature, { coercionType: "html" });
        }
      );
    } else {
      //Image is not embedded, or is referenced from template HTML
      Office.context.mailbox.item.body.setSignatureAsync(this.#signature_info.signature, { coercionType: "html" });
    }
  };
  /**
   * Gets template name (A,B,C) mapped based on the compose type
   * @param {*} compose_type The compose type (reply, forward, newMail)
   * @returns Name of the template to use for the compose type
   */
  #get_template_name = (compose_type) => {
    if (compose_type === "reply") return Office.context.roamingSettings.get("reply");
    if (compose_type === "forward") return Office.context.roamingSettings.get("forward");
    return Office.context.roamingSettings.get("newMail");
  };
  /**
   * Gets HTML signature in requested template format for given user
   * @param {"templateA" | "templateB" | "templateC"} template_name Which template format to use (A,B,C)
   * @returns {SignatureTemplate} HTML signature in requested template format
   */
  #get_signature_info = (template_name) => {
    /** @type {SignatureTemplate} */
    const returnSignature = {};
    /** @type {Signatures} */
    const signature = new Signatures(this.#user_info);
    if (template_name === "templateA") {
      const logoFileName = "sample-logo.png";
      returnSignature.signature = signature.get_template_A(logoFileName);
      returnSignature.logoBase64 = Common.defaultBase64Logo();
      returnSignature.logoFileName = logoFileName;
    }
    if (template_name === "templateB") {
      returnSignature.signature = signature.get_template_B();
    }
    if (template_name === "templateC") {
      returnSignature.signature = signature.get_template_C();
    }
    return returnSignature;
  };
}
/**
 * @typedef {Object} SignatureTemplate
 * @property {String} signature
 * @property {String} logoBase64
 * @property {String} logoFileName
 */
