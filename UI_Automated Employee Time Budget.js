const CustomUI = {
  titleMenu: "Workload Sheet Tools",
  updateSheet: "Update Activities Data from your calendar",
  createSheet: "Create New Worksheet",
  inputPrompt: {
    title: "Create New Worksheet",
    prompt: "Please enter your desire sheet name",
  },
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CustomUI.titleMenu)
    .addItem(CustomUI.createSheet, "menuCreateWorkloadSheet")
    .addItem(CustomUI.updateSheet, "menuUpdateWorkloadSheet")
    .addToUi();
}

function menuUpdateWorkloadSheet() {
  const ui = SpreadsheetApp.getUi();
  try {
    updateUserWorkloadSheet();
  } catch (e) {
    // Display an error message
    ui.alert(e.title, e.prompt, ui.ButtonSet.OK);
  }
}

function menuCreateWorkloadSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    CustomUI.inputPrompt.title,
    CustomUI.inputPrompt.prompt,
    ui.ButtonSet.OK_CANCEL
  );
  const button = response.getSelectedButton();
  const strInput = response.getResponseText();

  if (button === ui.Button.OK) {
    createNewWorkloadSheet(strInput);
    const shouldUpdate = ui.alert(
      "Update Activity",
      "Do you want to update your activity immediately?",
      ui.ButtonSet.YES_NO
    );

    if (shouldUpdate === ui.Button.YES) {
      menuUpdateWorkloadSheet();
    }
  }
}
