(function () {
  "use strict";

  // ─── Configuration ────────────────────────────────────────────────────────
  // When loaded via /embed.js this URL is injected automatically.
  // If you copy this file manually, replace __CHATBOT_URL__ with your server URL,
  // e.g. "https://chatbot.yourcompany.com"
  var CHATBOT_URL = "__CHATBOT_URL__";

  // ─── Inject iframe ────────────────────────────────────────────────────────
  function injectWidget() {
    if (document.getElementById("sp-chatbot-frame")) return;

    var frame = document.createElement("iframe");
    frame.id = "sp-chatbot-frame";
    frame.src = CHATBOT_URL;
    frame.title = "ON-PNT ChatBot";
    frame.setAttribute("allowtransparency", "true");
    frame.setAttribute("frameborder", "0");
    frame.setAttribute("scrolling", "no");

    Object.assign(frame.style, {
      position:      "fixed",
      bottom:        "0",
      right:         "0",
      width:         "420px",
      height:        "600px",
      border:        "none",
      background:    "transparent",
      zIndex:        "2147483647",
      pointerEvents: "none",
    });

    document.body.appendChild(frame);

    // ─── Listen for open/close signals from the widget ──────────────────────
    // The chatbot posts messages so the iframe pointer-events can be toggled,
    // allowing clicks to pass through the transparent areas to SharePoint.
    window.addEventListener("message", function (event) {
      if (event.origin !== CHATBOT_URL) return;

      if (event.data === "chatbot:open") {
        frame.style.pointerEvents = "all";
      } else if (event.data === "chatbot:close") {
        frame.style.pointerEvents = "none";
      }
    });
  }

  // Run after the DOM is ready
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", injectWidget);
  } else {
    injectWidget();
  }
})();
