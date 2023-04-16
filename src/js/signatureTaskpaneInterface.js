import Common from "./common";
import SignatureEventWrapper from "./signatureEventWrapper";
import SignatureSettingsInterface from "./signatureSettingsInterface";
import Signatures from "./signatures";

/* global document window Office */
export default class SignatureTaskpaneInterface {
  /** @type {import("./signatures").SignatureUserInfo} */
  #user_info = {};
  /**
   * Creates a new signature settings interface
   * @param {import("./signatures").SignatureUserInfo} user_info
   */
  constructor(user_info) {
    this.#user_info = user_info;
  }
  /**
   * Renders the interface in the DIV
   * @param {HTMLDivElement} container
   */
  render = (container) => {
    const html = `   
      <p class="templates">TEMPLATE A</p>
      <div id="templateA_box"></div>
      <button class="testBtn" id="buttonTemplateA">Test</button>
      <p class="templates">TEMPLATE B</p>
      <div id="templateB_box"></div>
      <button class="testBtn" id="buttonTemplateB">Test</button>
      <p class="templates">TEMPLATE C</p>
      <div id="templateC_box"></div>
      <button class="testBtn" id="buttonTemplateC">Test</button>
      <div style="height:35px;"><br/></div>
      <div class="dropdown">
        <label for="templates">New Mail
          <select style="margin-left:62px;" id="new_mail">
            <option value="none">None</option>
            <option value="templateA">Template A</option>
            <option value="templateB">Template B</option>
            <option value="templateC">Template C</option>
          </select>
        </label>
      </div>
      <div style="height:20px;"><br></div>
      <div class="dropdown">
        <label for="templates">Reply
            <select style="margin-left:85px;" id="reply">
              <option value="none">None</option>
              <option value="templateA">Template A</option>
              <option value="templateB">Template B</option>
              <option value="templateC">Template C</option>
            </select>
          </label>
      </div>
      <div style="height:20px;"><br></div>
      <div class="dropdown">
        <label for="templates">Forward
          <select style="margin-left:71px;" id="forward">
            <option value="none">None</option>
            <option value="templateA">Template A</option>
            <option value="templateB">Template B</option>
            <option value="templateC">Template C</option>
          </select>
        </label>
      </div>
      <div style="height:15px;"><br></div>
      <p id="newMApt"><strong>Appointments</strong> will use <strong>New Mail</strong> template.</p>
      <div style="height:5px;"><br></div>
      <p id="override_sig">Override Outlook signature</p>
      <label style = "position:relative; left:120px; top:15.5px;" class="switch">
        <input id="checkbox_sig" type="checkbox" checked>
        <span class="slider round"></span>
      </label>
      <div style="height:60px;"><br></div>
      <button id="edit_button" type="submit" class="registerbtn">Edit</button>
      <button id="submit_button" type="submit" class="registerbtn">Save</button>
      <p id="message">Success! You can close this pane.</p>`;
    container.innerHTML = html;
    /** @type {HTMLButtonElement} */
    const buttonTemplateA = document.getElementById("buttonTemplateA");
    buttonTemplateA.addEventListener("click", () => {
      const e = new SignatureEventWrapper(this.#user_info);
      e.addSignature("templateA");
    });
    /** @type {HTMLButtonElement} */
    const buttonTemplateB = document.getElementById("buttonTemplateB");
    buttonTemplateB.addEventListener("click", () => {
      const e = new SignatureEventWrapper(this.#user_info);
      e.addSignature("templateB");
    });
    /** @type {HTMLButtonElement} */
    const buttonTemplateC = document.getElementById("buttonTemplateC");
    buttonTemplateC.addEventListener("click", () => {
      const e = new SignatureEventWrapper(this.#user_info);
      e.addSignature("templateC");
    });

    /** @type {HTMLSelectElement} */
    const newMailInput = document.getElementById("new_mail");
    const newMailValue = Office.context.roamingSettings.get("newMail");
    if (Common.is_valid_data(newMailValue)) newMailInput.value = newMailValue;
    /** @type {HTMLSelectElement} */
    const forwardInput = document.getElementById("forward");
    const forwardValue = Office.context.roamingSettings.get("forward");
    if (Common.is_valid_data(forwardValue)) forwardInput.value = forwardValue;
    /** @type {HTMLSelectElement} */
    const replyInput = document.getElementById("reply");
    const replyValue = Office.context.roamingSettings.get("reply");
    if (Common.is_valid_data(replyValue)) replyInput.value = replyValue;
    /** @type {HTMLInputElement} */
    const checkSignature = document.getElementById("checkbox_sig");
    checkSignature.value = Office.context.roamingSettings.set("override");
    /** @type {HTMLButtonElement} */
    const editButton = document.getElementById("edit_button");
    editButton.addEventListener("click", () => {
      const settingsForm = new SignatureSettingsInterface();
      settingsForm.render(container);
    }); /** ADD HERE */
    /** @type {HTMLButtonElement} */
    const saveButton = document.getElementById("submit_button");
    saveButton.addEventListener("click", () => {
      let user_info_str = window.localStorage.getItem("user_info");
      if (user_info_str) {
        if (!this.#user_info) {
          this.#user_info = JSON.parse(user_info_str);
        }
      }
      Office.context.roamingSettings.set("user_info", user_info_str);
      Office.context.roamingSettings.set("newMail", newMailInput.selectedOptions[0].value);
      Office.context.roamingSettings.set("reply", replyInput.selectedOptions[0].value);
      Office.context.roamingSettings.set("forward", forwardInput.selectedOptions[0].value);
      Office.context.roamingSettings.set("override", checkSignature.checked);
      // save roaming settings
      Office.context.roamingSettings.saveAsync();
      if (checkSignature.checked === true) {
        Office.context.mailbox.item.disableClientSignatureAsync(() => {});
      }
    });
    const signature = new Signatures(this.#user_info);
    /**@type {HTMLDivElement} */
    const templateABox = document.getElementById("templateA_box");
    templateABox.innerHTML = signature.get_template_A();
    /**@type {HTMLDivElement} */
    const templateBBox = document.getElementById("templateB_box");
    templateBBox.innerHTML = signature.get_template_B();
    /**@type {HTMLDivElement} */
    const templateCBox = document.getElementById("templateC_box");
    templateCBox.innerHTML = signature.get_template_C();
  };
}
