(function () {
  "use strict";

  // ─── Configuration ────────────────────────────────────────────────────────
  // When loaded via /embed.js this URL is injected automatically.
  // If you copy this file manually, replace __CHATBOT_URL__ with your server URL,
  // e.g. "https://chatbot.yourcompany.com"
  var CHATBOT_URL = "__CHATBOT_URL__";

  // Iframe size when chat is closed (just big enough for the floating button)
  var CLOSED_W     = "100px";
  var CLOSED_H     = "100px";
  // Iframe size when chat is open (extra 40px each axis gives the box-shadow room)
  var OPEN_W       = "480px";
  var OPEN_H       = "660px";
  // Iframe size when chat is minimized (wide enough to show the pill)
  var MINIMIZED_W  = "320px";
  var MINIMIZED_H  = "80px";

  // ─── Inject iframe ────────────────────────────────────────────────────────
  function injectWidget() {
    if (document.getElementById("sp-chatbot-frame")) return;

    var frame = document.createElement("iframe");
    frame.id    = "sp-chatbot-frame";
    frame.src   = CHATBOT_URL;
    frame.title = "ON-PNT ChatBot";
    frame.setAttribute("allowtransparency", "true");
    frame.setAttribute("frameborder", "0");
    frame.setAttribute("scrolling", "no");

    Object.assign(frame.style, {
      position:   "fixed",
      bottom:     "0",
      right:      "0",
      width:      CLOSED_W,
      height:     CLOSED_H,
      border:     "none",
      background: "transparent",
      zIndex:     "2147483647",
      // pointer-events: all so the floating button inside is clickable.
      // The iframe starts small so it only covers the button area.
      pointerEvents: "all",
      overflow:   "hidden",
    });

    document.body.appendChild(frame);

    // ─── Resize on open/close signals from the widget ────────────────────
    var origin = new URL(CHATBOT_URL).origin;
    window.addEventListener("message", function (event) {
      if (event.origin !== origin) return;

      if (event.data === "chatbot:open") {
        frame.style.width  = OPEN_W;
        frame.style.height = OPEN_H;
      } else if (event.data === "chatbot:minimize") {
        frame.style.width  = MINIMIZED_W;
        frame.style.height = MINIMIZED_H;
      } else if (event.data === "chatbot:close") {
        frame.style.width  = CLOSED_W;
        frame.style.height = CLOSED_H;
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
