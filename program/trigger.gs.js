/* НЕАКТИВЕН
 * Орегинальный файл - незадействован
 * сохранение ярлыка в письме? Заменен файлом triggerFromSheet.gs.js
 * @param e
 * @returns {*}
 */
function trigger_SaveEmails(e) {

	try {

		var lock = LockService.getScriptLock();

		lock.tryLock(1000 * 10);

		if (!lock.hasLock()) {
			writeLog_("[Error] Could not acquire lock, server is busy");
			return "Another instance of the script is currently running. Please try after few minutes.";
		}

		var start = new Date(),
			archiveName = "Saved",
			archive = getGmailLabel_(archiveName),
			count = 0,
			minutes = 4,
			batchsize = 50,
			rules = getRules_();

		if (e && e.ruleID) {
			rules = getRules_(e.ruleID);
			minutes = 3;
		}
		if (e && e.batchSize) {
			batchsize = e.batchSize;
		}

		for (var key in rules) {

			try {
				var folder = DriveApp.getFolderById(rules[key].action.savefolderID);
			} catch(f) {
				writeLog_("[Error] Folder ID #" + rules[key].action.savefolderID + " " + f.toString());
				continue;
			}

			var filter = rules[key].rule,
				savepdf = rules[key].action.saveemail,
				saveatt = rules[key].action.savefiles,
				threads = GmailApp.search(filter + " -label:" + archiveName, 0, batchsize);

			if (folder && threads.length) {

				writeLog_("[Rule] " + filter + " [" + threads.length + " threads]");

				for (var x = 0; x < threads.length; x++) {

					if (isTimeUp_(start, minutes)) {
						break;
					}

					count++;

					threads[x].addLabel(archive);

					var ids = [],
						html = "",
						messages = threads[x].getMessages();

					for (var m = 0; m < messages.length; m++) {

						var message = messages[m],
							id = message.getId();

						if (fileExists_(id)) {
							continue;
						}

						var file, files = [],
							att = message.getAttachments(),
							subject = message.getSubject(),
							date = formatDate(message);

						ids.push(id);

						if (saveatt) {

							for (var z = 0; z < att.length; z++) {
								try {
									file = folder.createFile(att[z]).setDescription([id, subject, date].join("\n\n"));
									files.push(file);
									writeLog_("[File] " + file.getName(), file.getUrl());
								} catch (f) {
									writeLog_("[Error] Saving File #" + id + " " + f.toString());
								}
							}

						}

						if (savepdf) {

							var from = formatEmails_(message.getFrom()),
								to = formatEmails_(message.getTo()),
								cc = formatEmails_(message.getCc()),
								body = message.getBody(),
								raw = message.getRawContent().replace(/=\r\n([^-][^-])/g, "$1").replace(/=3D/g, "=");

							if (cc !== "&nbsp;") {
								cc = '<dt>Cc:</dt> <dd>' + cc + '</dd>\n';
							} else {
								cc = "";
							}

							html += '<dl class="email-meta">\n' +
								'<dt>From:</dt><dd class="avatar" style="background:' + getBackgroundColor_() + '">' + getLetter_(from) + '</dd><dd class="strong">' + from + '</dd>\n' +
								'<dt>Subject:</dt> <dd>' + subject + '</dd>\n' +
								'<dt>Date:</dt> <dd>' + date + '</dd>\n' +
								'<dt>To:</dt> <dd>' + to + '</dd>\n' + cc +
								'</dl>\n';

							try {
								body = embedHtmlImages_(body);
								body = embedInlineImages_(body, raw);
							} catch (b) {}

							html += body;

							if (files.length > 0) {

								html += '<br />\n<strong>Attachments:</strong>\n' +
									'<div class="email-attachments">\n';

								for (var f in files) {
									html += '<a href="' + files[f].getUrl() + '">' + files[f].getName() + '</a> ' + humanFileSize_(files[f].getSize()) + '<br>\n';
								}

								html += '</div>\n';

							}

						}

					}


					if (savepdf && (html !== "")) {

						html = '<html>\n' +
							'<style type="text/css">\n' +
							'body{padding:0 10px;min-width:700px;-webkit-print-color-adjust: exact;}' +
							'body>dl.email-meta{font-family:"Helvetica Neue",Helvetica,Arial,sans-serif;font-size:14px;padding:0 0 10px;margin:0 0 5px;border-bottom:1px solid #ddd;page-break-before:always}' +
							'body>dl.email-meta:first-child{page-break-before:auto}' +
							'body>dl.email-meta dt{color:#808080;float:left;width:60px;clear:left;text-align:right;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-style:normal;font-weight:700;line-height:1.4}' +
							'body>dl.email-meta dd{margin-left:70px;line-height:1.4}' +
							'body>dl.email-meta dd a{color:#808080;font-size:0.85em;text-decoration:none;font-weight:normal}' +
							'body>dl.email-meta dd.avatar{float: right;background: lightgreen;width: 72px;height: 72px;border-radius: 36px;color: white;text-align:center;font-size:36px;line-height:72px;}' +
							'body>dl.email-meta dd.avatar img{max-height:72px;max-width:72px;border-radius:36px}' +
							'body>dl.email-meta dd.strong{font-weight:bold}' +
							'body>div.email-attachments{font-size:0.85em;color:#999}' +
							'</style>\n' +
							'<body>\n' + html + '\n</body>\n</html>';

						var fileName = "[Email] " + sanitizeFilename_(threads[x].getFirstMessageSubject()) + ".pdf";
						var blob = Utilities.newBlob(html, 'text/html');
						var pdf = folder.createFile(blob.getAs('application/pdf'))
														.setName(fileName)
														.setDescription(ids.join("\n"));

						writeLog_(fileName, pdf.getUrl());

					}

				}
			}
		}
	} catch (f) {
		writeLog_("[Error] " + f.toString());
	}

	lock.releaseLock();

	return count;
}