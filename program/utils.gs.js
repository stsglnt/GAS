/*
 * Запись логов в поля таблици
 */
function writeLog_() {

    try {

        var row = [Utilities.formatDate(new Date(), TIMEZONE, "MMM-dd HH:mm:ss")];

        for (var i=0; i < arguments.length; i++) {
            row.push(arguments[i]);
        }

        SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].appendRow(row);

    } catch (f) {}
}

/*
 * TIMEZONE
 * @returns {string}
 */
function getTZ_() {
    return SpreadsheetApp.getActive().getSpreadsheetTimeZone();
}

function fileExists_(id) {
    try {
        var files = DriveApp.searchFiles('fullText contains "' + id + '"');
        return files.hasNext() ? true : false;
    } catch (f) {}
    return false;
}

/*
 * перевод байтов
 * @param size
 * @returns {*}
 */
function humanFileSize_(size) {
    try {
        if (isNaN(size) || size == 0) return "";
        var i = Math.floor(Math.log(size) / Math.log(1024));
        return ' (' + (size / Math.pow(1024, i)).toFixed(2) * 1 + ' ' + ['B', 'kB', 'MB', 'GB', 'TB'][i] + ')';
    } catch (f) {}
    return "";
}

/*
 * рандомный цвет фона
 * @returns {string}
 */
function getBackgroundColor_() {
    var colors = ["#2ecc71", "#3498db", "#34495e", "#e74c3c", "#16a085", "#f1c40f", "#7f8c8d", "#c0392b", "#2c3e50"];
    var index = Math.floor(Math.random() * colors.length);
    return colors[index];
}

/*
 * Первая буква заглавная
 * @param str
 * @returns {string}
 */
function getLetter_(str) {
    str = str.match(/[A-Za-z0-9]/);
    return str ? str[0].toUpperCase() : "!";
}

/*
 * создание и удаление "тригера"  метки?
 * @param enableTrigger {boolean}
 */
function toggleTrigger_(enableTrigger) {

    var properties = getProps_(); // Сессия ?
    var triggerId = properties.getProperty('ctrlqSaveEmailsTrigger'); // Сессия ?

    if (!enableTrigger && triggerId != null) {
        /* Удаление !enableTrigger === true */

        var triggers = ScriptApp.getProjectTriggers();
        for (var i = 0; i < triggers.length; i++) {
            if (triggers[i].getUniqueId() == triggerId) {
                ScriptApp.deleteTrigger(triggers[i]);
                writeLog_("[Trigger] Deleted");
                break;
            }
        }
        properties.deleteProperty('ctrlqSaveEmailsTrigger');

    } else if (enableTrigger && triggerId == null) {

        /* Сохранение/создание enableTrigger === false */
        var trigger = ScriptApp.newTrigger('trigger_SaveEmails_sheet').timeBased().everyMinutes(15).create();
        writeLog_("[Trigger] Created");
        properties.setProperty('ctrlqSaveEmailsTrigger', trigger.getUniqueId());

    }

    onOpen();

}

/*
 * PropertiesService.getUserProperties - Позволяет сценариям хранить простые данные в парах ключ-значение,
 * привязанных к одному сценарию, Свойства не могут быть разделены между сценариями.
 * @returns {Properties}
 * @private
 */
function getProps_() {
    return PropertiesService.getUserProperties();
}

/*
 * НЕАКТИВНОЕ
 * сброс/удаление всех правил ?
 */
function reset() {
    toggleTrigger_(false);
    getProps_().deleteAllProperties();
    SpreadsheetApp.getActive().toast("All rules have been deleted.");
}

/*
 * НЕАКТИВНОЕ
 */
function isTriggerActive_() {
    return getProps_().getProperty('ctrlqSaveEmailsTrigger') ? true : false;
}

/*
 * НЕАКТИВНОЕ
 */
function writeLogs_(log) {
    if (log !== "") {
        getProps_().setProperty("ctrlqSaveEmailsLog", log);
    }
}

/*
 * НЕАКТИВНОЕ
 * отправка багов разработчику
 */
function emailLogs() {
    MailApp.sendEmail("amit@labnol.org", "[Save Emails] Log for " + getUserEmail_(), JSON.stringify(getProps_().getProperties()));
    SpreadsheetApp.getActive().toast("The debug logs were emailed to amit@labnol.org");
}

/*
 * НЕАКТИВНОЕ
 *  For File Picker
 * @returns {string}
 */
function getOAuthToken() {
    DriveApp.getRootFolder();
    return ScriptApp.getOAuthToken();
}

/* время ожидания работы
 * Prevent Timeout
 * @param start начало отсчета времени
 * @param minutes время в минутах
 * @returns {boolean}
 */
function isTimeUp_(start, minutes) {
    var now = new Date();
    return now.getTime() - start.getTime() > minutes*1000*60; // 4 minutes
}

/*
 * получаем метки с почтового адреса
 * @param str
 * @returns {GmailLabel}
 */
function getGmailLabel_(str) {
    var label = GmailApp.getUserLabelByName(str);
    return label ? label : GmailApp.createLabel(str);
}

/*
 * получение мейла активного пользователя
 * @returns {string}
 */
function getUserEmail_() {
    var email = Session.getActiveUser().getEmail();
    if (email === "") {
        email = Session.getEffectiveUser().getEmail();
    }
    return email;
}

/*
 * НЕАКТИВНОЕ
 * нормализатор текста
 * @param str
 * @returns {string}
 */
function normalize_(str) {
    return str.replace(/[^\w]+/g, "").toLowerCase();
}

/*
 * Запуск скрипта вручную с окна выбора правил
 * @param ruleID
 * @returns {*}
 */
function runRule(ruleID) {

    try {
        var rule = getRulesFromSheet(ruleID);
        var result = trigger_SaveEmails_sheet({rule: rule[0], ruleID: ruleID, batchSize: 20});

        if (isNaN(result)) {
            return result;
        }

        SpreadsheetApp.getActive().toast("Rule processed successfully.");

        return "Rule processed. " + result + " emails were added to your <a href='https://drive.google.com/drive/recent' target='_blank'>Google Drive</a>.";

    } catch (f) {}

    return "Sorry, we had trouble running this rule";
}

/*
 * НЕАКТИВНОЕ
 * получаем метки с почтового адреса
 * @returns {[string,string,string]} масив из строк-меток
 */
function getGmailLabels_() {

    var all = ["Inbox", "Starred", "Important"];

    var labels = GmailApp.getUserLabels();
    for (var l in labels) {
        all.push(labels[l].getName());
    }

    all.push("All");
    all.push("Spam");
    all.push("Trash");

    return all;

}

/*
 * Поиск существуюших правил ?
 * @param id
 * @returns {string|null|string}
 */
function getRules_(id) {
    var props = getProps_(); // сеанс пользователя ?
    var rules = props.getProperty("ctrlqSaveEmailRules") || "{}";

    try {
        rules = JSON.parse(rules);
    } catch (e) {
        rules = {};
    }
    javascript:;
    if (id) {
        for (var key in rules) {
            if (key !== id) {
                delete rules[key];
            }
        }
    }

    return rules;
}

/*
 * Поис сушествуюших правил ?
 * @returns {number}
 */
function rulesCount() {
    var rules = getRules_();
    Logger.log(rules);
    var count = 0;
    for (var key in rules) {
        count++;
    }
    return count;
}

/*
 * fixed rule ID
 * @param id
 * @returns {string[]}
 */
function getRulesFromSheet(id) {
    var objDB = getObjDB();
    var rules = objDB.getRows(ssDB, 'rules');

    rules.forEach(function(rule) {
        if(!rule.ruleID) {
            rule['ruleID'] = md5_(rule.rule);
            objDB.updateRow(ssDB, 'rules', {ruleID: md5_(rule.rule) }, {rule: rule.rule, savefolderID: rule.savefolderID})
        }
    });

    if (id) {
        rules = objDB.getRows( ssDB, 'rules', [], {ruleID: String(id)});
    }

    return rules
}

function updateTrigger_() {
    toggleTrigger_(rulesCount() > 0 ? true : false);
}

/*
 * НЕАКТИВНОЕ
 * Сохранение правил
 * @param e
 */
function saveRule(e) {

    try {

        var rule = e.rule.trim();

        if (rule !== "") {

            var props = getProps_();
            var rules = getRules_();
            var ruleID = md5_(rule);

            rules[ruleID] = e;
            e.action['ruleID'] = ruleID;

            props.setProperty("ctrlqSaveEmailRules", JSON.stringify(rules));
            addRuleToTable(e.action);
            writeLog_("[Save] Rule " + JSON.stringify(e));

            updateTrigger_();

            var time_save_emails = 15;
            trigger_SaveEmails_sheet({rule: e.action, ruleID: ruleID, batchSize: 1});
            if (e.action.isactive){
                SpreadsheetApp.getActive().toast("It will automatically save matching emails to your Google Drive every "+time_save_emails+" minutes.", "Rule Created", time_save_emails);
            } else {
                SpreadsheetApp.getActive().toast("Set as not active", "Rule Created", time_save_emails);
            }

            return;

        }

    } catch (f) {writeLog_("[Error] " + f.toString());}

    SpreadsheetApp.getActive().toast("Sorry, we had trouble creating this rule. Try later");

}

/*
 * НЕАКТИВНОЕ
 * Delete rule
 * @param ruleID selected rule
 * @returns {*}
 */
function deleteRule(ruleID) {

    try {

        var props = getProps_();
        var rules = getRules_();

        for (var key in rules) {
            if (key === ruleID) {
                delete rules[key];
                props.setProperty("ctrlqSaveEmailRules", JSON.stringify(rules));
                break;
            }
        }

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        ssDB = objDB.open( ss.getId());
        objDB.deleteRow( ssDB, 'rules', {ruleID:ruleID });

        updateTrigger_();

        writeLog_("[Delete] Rule " + ruleID + " deleted");

        SpreadsheetApp.getActive().toast("Rule successfully deleted");
        return ruleID;

    } catch (f) { writeLog_("[Error] " + f.toString());}

    return "Sorry, we had trouble deleting the rule";
}

/*
 * append row with rules
 * @param newRule
 */
function addRuleToTable(newRule) {
    var objDB = getObjDB();
    var rowCount = objDB.insertRow(ssDB, 'rules', newRule);
    Logger.log( rowCount );
}

/**
 * get objDB
 * @returns {*}
 */
function getObjDB() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ssDB = objDB.open( ss.getId());
    objDB.setSkipRows(ssDB, 'rules', 1, 1);
    return objDB
}