/// <reference path="../../node_modules/@types/office-js/index.d.ts" />
import Common from "./common";
import SignatureTaskpaneInterface from "./signatureTaskpaneInterface";
/* global Office document window */
export default class SignatureSettingsInterface {
  /** @type {import("./signatures").SignatureUserInfo} */
  #user_info = {};
  /** @type {HTMLElement} */
  #message = null;
  /**
   * Renders the interface in the container div
   * @param {HTMLElement} container
   */
  render = (container) => {
    const html = `
      <h2>Add signature data</h2>
      <input type="text" id="display_name" placeholder="Name*" required />
      <input type="email" id="email_id" placeholder="Email address*" required />
      <input type="text" id="job_title" placeholder="Title" />
      <input type="text" id="phone_number" placeholder="Phone number" />
      <input type="text" placeholder="Eg: Thank you," id="greeting_text" />
      <input type="text" placeholder="Eg: She, Her" id="preferred_pronoun" />
      <button id="next_button_t1" class="registerbtn">Save</button>
      <button id="reset_all_config_btn" class="registerbtn">RESET ALL</button>
      <p id="message"></p>
    `;
    container.innerHTML = html;
    /** @type {HTMLInputElement} */
    const displayName = document.getElementById("display_name");
    displayName.addEventListener("click", () => displayName.select());
    /** @type {HTMLInputElement} */
    const emailId = document.getElementById("email_id");
    emailId.addEventListener("click", () => emailId.select());
    /** @type {HTMLInputElement} */
    const jobTitle = document.getElementById("job_title");
    /** @type {HTMLInputElement} */
    const phoneNumber = document.getElementById("phone_number");
    /** @type {HTMLInputElement} */
    const greeting = document.getElementById("greeting_text");
    /** @type {HTMLInputElement} */
    const pronoun = document.getElementById("preferred_pronoun");
    /** @type {HTMLButtonElement} */
    const nextButton = document.getElementById("next_button_t1");
    nextButton.addEventListener("click", () => {
      let name = displayName.value.trim();
      let email = emailId.value.trim();
      this.#display_message(null);
      if (this.#validate_form(name, email)) {
        this.#display_message(null);
        this.#user_info.name = name;
        this.#user_info.email = email;
        this.#user_info.job = jobTitle.value.trim();
        this.#user_info.phone = phoneNumber.value.trim();
        this.#user_info.greeting = greeting.value.trim();
        this.#user_info.pronoun = pronoun.value.trim();
        if (this.#user_info.pronoun !== "") {
          this.#user_info.pronoun = "(" + this.#user_info.pronoun + ")";
        }
        window.localStorage.setItem("user_info", JSON.stringify(this.#user_info));
        const taskPane = new SignatureTaskpaneInterface(this.#user_info);
        taskPane.render(container);
      }
    });
    /** @type {HTMLButtonElement} */
    const resetButton = document.getElementById("reset_all_config_btn");
    resetButton.addEventListener("click", () => {
      // clear form
      displayName.value = "";
      emailId.value = "";
      jobTitle.value = "";
      phoneNumber.value = "";
      greeting.value = "";
      pronoun.value = "";
      // clear storage
      window.localStorage.removeItem("user_info");
      window.localStorage.removeItem("newMail");
      window.localStorage.removeItem("reply");
      window.localStorage.removeItem("forward");
      window.localStorage.removeItem("override");
      // clear office settings
      Office.context.roamingSettings.remove("user_info");
      Office.context.roamingSettings.remove("newMail");
      Office.context.roamingSettings.remove("reply");
      Office.context.roamingSettings.remove("forward");
      Office.context.roamingSettings.remove("override");
      Office.context.roamingSettings.saveAsync(
        /** @param {Office.AsyncResult} asyncResult */
        (asyncResult) => {
          let message =
            "All settings reset successfully! This add-in won't insert any signatures. You can close this pane now.";
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            message = "Failed to reset. Please try again.";
          }
          this.#display_message(message);
        }
      );
    });
    this.#message = document.getElementById("message");
    // load defaults - just in case
    displayName.value = Office.context.mailbox.userProfile.displayName;
    emailId.value = Office.context.mailbox.userProfile.emailAddress;
    // load the user settings
    this.#load_saved_user_info();
    displayName.value = this.#user_info.name ? this.#user_info.name : displayName.value;
    emailId.value = this.#user_info.email ? this.#user_info.email : emailId.value;
    jobTitle.value = this.#user_info.job ? this.#user_info.job : "";
    phoneNumber.value = this.#user_info.phone ? this.#user_info.phone : "";
    greeting.value = this.#user_info.greeting ? this.#user_info.greeting : "";
    let p = this.#user_info.pronoun;
    if (p && p.length >= 3) {
      pronoun.value = p.substring(1, p.length - 1);
    } else if (p) {
      pronoun.value = p;
    } else {
      pronoun.value = "";
    }
  };
  /**
   * Loads the users saved data
   */
  #load_saved_user_info = () => {
    let user_info_str = window.localStorage.getItem("user_info");
    if (!user_info_str) {
      user_info_str = Office.context.roamingSettings.get("user_info");
    }
    if (user_info_str) {
      this.#user_info = JSON.parse(user_info_str);
    }
  };
  /**
   * Displays the message to the user
   * @param {String} [msg]
   */
  #display_message = (msg = null) => {
    if (msg === null) this.#message.innerText = "";
    else this.#message.innerText = msg;
  };
  /**
   * Validates the form data
   * @param {String} name
   * @param {String} email
   * @returns {Boolean}
   */
  #validate_form = (name, email) => {
    if (!Common.is_valid_data(name)) {
      this.#display_message("Please enter a valid name.");
      return false;
    }
    if (!Common.is_valid_email(email)) {
      this.#display_message("Please enter a valid email address.");
      return false;
    }
    return true;
  };
}
