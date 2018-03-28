function categorize() {　　　　
	Browser.msgBox("予定のカテゴライズを開始します。");

	var categorySheet = SpreadsheetApp.getActive().getSheetByName("カテゴリ設定");　　　　
	var categories　 = categorySheet.getRange("A2:C50").getValues();

	var scheduleSheet = SpreadsheetApp.getActive().getSheetByName("スケジュール");　　

	var schedules = scheduleSheet.getRange("B3:O50").getValues();

	var word;
	var range;
	var headerColor;

	for　 (var i　 = 　0; i　 < 　schedules.length; i++)　 {　　　　　　　　　　　　　　
		for　 (var j　 = 　0; j　 < schedules[i].length; j++)　 {　
			range = scheduleSheet.getRange(i　 + 　3, j　 + 　2);

			for　 (var k　 = 　0; k　 < 　categories.length; k++)　 {
				if (categories[k][0] != "") {
					word = schedules[i][j];

					if (word.indexOf("\r\n") <= -1 &&
						word.indexOf(categories[k][0]) > -1) {
						range.setValue(categories[k][1]);
						range.setBackground(categorySheet.getRange(k + 2, 3).getBackground());
					}
				}
			}

			headerColor = scheduleSheet.getRange(1, j　 + 　2).getBackground();

			if (headerColor == "#e06666" || headerColor == "#6d9eeb") {
				range.setBackground(headerColor);
			}
		}　　　　　　　　　　　　
	}　　　　　　　

	Browser.msgBox("予定のカテゴライズが完了しました。");
}
