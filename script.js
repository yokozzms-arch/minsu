(function () {
  "use strict";

  var chatBody = document.getElementById("chatBody");
  var input = document.getElementById("messageInput");
  var sendBtn = document.getElementById("sendBtn");

  // 현재 시각을 "오전/오후 h:mm" 형식으로 반환
  function currentTime() {
    var now = new Date();
    var h = now.getHours();
    var m = now.getMinutes();
    var ampm = h < 12 ? "오전" : "오후";
    var h12 = h % 12;
    if (h12 === 0) h12 = 12;
    var mm = m < 10 ? "0" + m : "" + m;
    return ampm + " " + h12 + ":" + mm;
  }

  // 안전한 텍스트 노드 생성 (XSS 방지)
  function textEl(tag, className, text) {
    var el = document.createElement(tag);
    if (className) el.className = className;
    if (text != null) el.textContent = text;
    return el;
  }

  // 메시지 말풍선 추가
  // opts: { text, mine, name, avatar, unread }
  function addMessage(opts) {
    var row = document.createElement("div");
    row.className = "msg-row " + (opts.mine ? "me" : "other");

    // 상대방이면 프로필
    if (!opts.mine) {
      var pic = textEl("div", "profile-pic", null);
      pic.textContent = opts.avatar || "🙂";
      row.appendChild(pic);
    }

    var content = textEl("div", "msg-content", null);

    // 상대방 이름
    if (!opts.mine && opts.name) {
      content.appendChild(textEl("div", "sender-name", opts.name));
    }

    var line = textEl("div", "bubble-line", null);
    var bubble = textEl("div", "bubble", opts.text);

    var meta = document.createElement("span");
    if (opts.mine && opts.unread) {
      meta.className = "msg-unread";
      meta.textContent = "1";
      line.appendChild(meta);
    }

    line.appendChild(bubble);
    line.appendChild(textEl("span", "msg-time", currentTime()));

    content.appendChild(line);
    row.appendChild(content);
    chatBody.appendChild(row);

    scrollToBottom();
    return bubble;
  }

  function scrollToBottom() {
    chatBody.scrollTop = chatBody.scrollHeight;
  }

  // 자동 응답용 예시 문구
  var autoReplies = [
    "ㅇㅇ 알겠어!",
    "오 좋다 👍",
    "지금 확인해볼게~",
    "ㅋㅋㅋㅋ 진짜?",
    "그래 그러자!",
    "잠깐만, 곧 갈게",
  ];
  var replyIndex = 0;

  function nextReply() {
    var r = autoReplies[replyIndex % autoReplies.length];
    replyIndex++;
    return r;
  }

  // 메시지 전송
  function send() {
    var text = input.value.trim();
    if (!text) return;

    addMessage({ text: text, mine: true, unread: true });

    input.value = "";
    autoResize();
    updateSendBtn();
    input.focus();

    // 상대방 자동 응답 (데모용 로컬 동작)
    setTimeout(function () {
      addMessage({
        text: nextReply(),
        mine: false,
        name: "민수",
        avatar: "😎",
      });
    }, 700 + Math.floor(Math.random() * 600));
  }

  // textarea 높이 자동 조절
  function autoResize() {
    input.style.height = "auto";
    input.style.height = Math.min(input.scrollHeight, 100) + "px";
  }

  // 전송 버튼 활성/비활성
  function updateSendBtn() {
    sendBtn.disabled = input.value.trim().length === 0;
  }

  // 이벤트 바인딩
  sendBtn.addEventListener("click", send);

  input.addEventListener("input", function () {
    autoResize();
    updateSendBtn();
  });

  input.addEventListener("keydown", function (e) {
    // Enter 전송, Shift+Enter 줄바꿈
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      send();
    }
  });

  // 초기 데모 메시지
  window.addEventListener("DOMContentLoaded", function () {
    addMessage({
      text: "안녕! 카카오톡 스타일 채팅이야 🎉",
      mine: false,
      name: "민수",
      avatar: "😎",
    });
    addMessage({
      text: "아래에 메시지를 입력해서 보내봐!",
      mine: false,
      name: "민수",
      avatar: "😎",
    });
    updateSendBtn();
    input.focus();
  });

  updateSendBtn();
})();
