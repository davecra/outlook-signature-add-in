/// <reference path="../../node_modules/@types/office-js/index.d.ts" />
/* global Office */
export default class Common {
  /**
   * Validates the string
   * @param {String} str
   * @returns {Boolean}
   */
  static is_valid_data = (str) => {
    return str !== null && str !== undefined && str !== "" && str.length > 0;
  };
  /**
   * Validates an email address
   * @param {String} email
   * @returns {Boolean}
   */
  static is_valid_email = (email) => {
    /** @type {RegExp} */
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return this.is_valid_data(email) && emailRegex.test(email);
  };
  /**
   * Returns breaks for calendar pos
   * @returns {String}
   */
  static get_cal_offset = () => {
    return "<br/><br/>";
  };
  /**
   * Creates information bar to display when new message or appointment is created
   */
  static display_insight_infobar = () => {
    Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
      type: "insightMessage",
      message: "Please set your signature with the Office Add-ins sample.",
      icon: "Icon.16x16",
      actions: [
        {
          actionType: "showTaskPane",
          actionText: "Set signatures",
          commandId: this.#get_command_id(),
          contextData: "{''}",
        },
      ],
    });
  };
  /**
   * Gets correct command id to match to item type (appointment or message)
   * @returns The command id
   */
  static #get_command_id = () => {
    if (Office.context.mailbox.item.itemType == "appointment") {
      return "MRCS_TpBtn1";
    }
    return "MRCS_TpBtn0";
  };
  /**
   * Returns the default logo
   * @returns {String}
   */
  static defaultBase64Logo = () => {
    return "iVBORw0KGgoAAAANSUhEUgAAACIAAAAiCAYAAAA6RwvCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAEeSURBVFhHzdhBEoIwDIVh4EoeQJd6YrceQM+kvo5hQNokLymO/4aF0/ajlBl1fL4bEp0uj3K9XQ/lGi0MEcB3UdD0uVK1EEj7TIuGeBaKYCgIswCLcUMid8mMcUEiCMk71oRYE+Etsd4UD0aFeBBSFtOEMAgpg6lCIggpitlAMggpgllBeiAkFjNDeiIkBlMgeyAkL6Z6WJdlEJJnjvF4vje/BvRALNN23tyRXzVpd22dHSZtLhjMHemB8cxRINZZyGCssbL2vCN7YLwItHo0PTEMAm3OSA8Mi0DVw5rBRBCoCkERTBSBmhDEYDII5PqlZy1iZSGQuiOSZ6JW3rEuCIpgmDFuCGImZuEUBHkWiOweDUHaQhEE+pM/aobhBZaOpYLJeeeoAAAAAElFTkSuQmCC";
  };
}
