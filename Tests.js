function testManageProperty() {
  
  // Set test flag
  TEST = true;

  manageProperty();

}

function testGetBills() {
   
  // Set test flag
  TEST = true;

  getBills();

}

function testEditTrigger() {
  
  let e = {
    range: LEDGER.getRange(LEDGER.getDataRange().getNumRows(),4)
  }
  onEdit(e);
}

function apiTest(message) {
  return 'Hello ' + message;
}

function testAddTask() {
  let id = '185eeb3794bb9001';
  addTask(id, 'test task');
}