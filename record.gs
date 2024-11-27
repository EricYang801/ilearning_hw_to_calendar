function checkEmailsAndAddToCalendar() {
  var label = GmailApp.getUserLabelByName("iLearning");  // 取出加了「iLearning」標籤的郵件
  var threads = label.getThreads();  // 取得所有該標籤下的郵件
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();  // 取得每個線程中的所有郵件
    
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      
      if (message.isUnread()) {  // 只處理未讀郵件
        Logger.log(message.getSubject());  // 打印郵件的標題
        
        // 解析郵件內容
        var subject = message.getSubject();
        var body = message.getBody();
        
        // 提取課程名稱
        var courseNameMatch = subject.match(/\[課程: (.+?)\]/) || body.match(/課程名稱[：: ]*(.+)/);
        
        // 提取作業標題
        var assignmentNameMatch = body.match(/作業名稱[：: ]*(.+)/) || subject.match(/\[作業\]「(.+?)」/);
        
        // 提取繳交期限
        var deadlineMatch = body.match(/繳交期限[：: ]*(\d{4}-\d{2}-\d{2} \d{2}:\d{2})/) || body.match(/繳交期限[：: ]*(\d{4}-\d{2}-\d{2})/);
        
        // 確保找到所有需要的資訊
        if (courseNameMatch && deadlineMatch) {
          var courseName = courseNameMatch[1].trim() || "未指定課程"; // 去除空白
          var assignmentName = assignmentNameMatch ? assignmentNameMatch[1].trim() : "未指定作業"; // 去除空白
          var deadline = new Date(deadlineMatch[1]);
          
          // 將作業資訊添加到 Google Calendar
          var calendar = CalendarApp.getDefaultCalendar();
          calendar.createEvent(
            courseName + ' - ' + assignmentName, 
            new Date(),  // 設定當前時間為開始時間
            deadline,  // 設定繳交期限
            {description: '作業繳交截止'}
          );
          
          message.markRead();  // 標記郵件為已讀
        } else {
          Logger.log("郵件內容無法正確解析");
        }
      }
    }
  }
}
