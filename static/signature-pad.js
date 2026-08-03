/* Small local signature pad: no external CDN or third-party runtime. */
(function (global) {
  "use strict";
  function SignaturePad(canvas, options) {
    this.canvas = canvas;
    this.context = canvas.getContext("2d");
    this.options = options || {};
    this.empty = true;
    this.drawing = false;
    this.last = null;
    this.context.lineCap = "round";
    this.context.lineJoin = "round";
    this.context.strokeStyle = this.options.penColor || "#111827";
    this.context.lineWidth = this.options.minWidth || 1.5;
    var self = this;
    function point(event) {
      var touch = event.touches && event.touches[0];
      var source = touch || event;
      var rect = canvas.getBoundingClientRect();
      return {x: (source.clientX - rect.left) * canvas.width / rect.width,
              y: (source.clientY - rect.top) * canvas.height / rect.height};
    }
    function start(event) {
      event.preventDefault(); self.drawing = true; self.last = point(event);
    }
    function move(event) {
      if (!self.drawing) return;
      event.preventDefault(); var next = point(event); var ctx = self.context;
      ctx.beginPath(); ctx.moveTo(self.last.x, self.last.y); ctx.lineTo(next.x, next.y); ctx.stroke();
      self.last = next; self.empty = false;
    }
    function end(event) { if (event) event.preventDefault(); self.drawing = false; self.last = null; }
    canvas.addEventListener("pointerdown", start, {passive:false});
    canvas.addEventListener("pointermove", move, {passive:false});
    global.addEventListener("pointerup", end, {passive:false});
    canvas.addEventListener("touchstart", start, {passive:false});
    canvas.addEventListener("touchmove", move, {passive:false});
    global.addEventListener("touchend", end, {passive:false});
  }
  SignaturePad.prototype.clear = function () {
    this.context.clearRect(0, 0, this.canvas.width, this.canvas.height); this.empty = true;
  };
  SignaturePad.prototype.isEmpty = function () { return this.empty; };
  SignaturePad.prototype.toDataURL = function (type) { return this.canvas.toDataURL(type || "image/png"); };
  global.SignaturePad = SignaturePad;
})(window);
