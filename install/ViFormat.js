// Hop thoai co thanh cuon (Unicode) cho ViFormat tren Mac. Goi: osascript -l JavaScript ViFormat.js <file.txt> [auto]
// Dong bat dau bang "#" la tieu de (in dam).
ObjC.import('Cocoa');
function run(argv) {
  var raw = $.NSString.stringWithContentsOfFileEncodingError($(argv[0]), $.NSUTF8StringEncoding, null);
  var lines = ObjC.unwrap(raw).split('\n');
  var app = $.NSApplication.sharedApplication;
  app.setActivationPolicy($.NSApplicationActivationPolicyAccessory);
  app.activateIgnoringOtherApps(true);

  var normal = $.NSFont.systemFontOfSize(13), bold = $.NSFont.boldSystemFontOfSize(14);
  var out = $.NSMutableAttributedString.alloc.init;
  lines.forEach(function (l) {
    var head = l.charAt(0) === '#';
    var s = (head ? l.slice(1) : l) + '\n';
    var piece = $.NSMutableAttributedString.alloc.init;
    piece.replaceCharactersInRangeWithString($.NSMakeRange(0, 0), $(s));
    var rg = $.NSMakeRange(0, ObjC.unwrap($(s)).length);
    piece.addAttributeValueRange($.NSFontAttributeName, head ? bold : normal, rg);
    piece.addAttributeValueRange($.NSForegroundColorAttributeName, $.NSColor.labelColor, rg);
    out.appendAttributedString(piece);
  });

  var W = 600, H = 420;
  var sv = $.NSScrollView.alloc.initWithFrame($.NSMakeRect(0, 0, W, H));
  sv.hasVerticalScroller = true;
  sv.borderType = $.NSBezelBorder;
  var tv = $.NSTextView.alloc.initWithFrame($.NSMakeRect(0, 0, W - 20, H));
  tv.editable = false;
  tv.verticallyResizable = true;
  tv.autoresizingMask = $.NSViewWidthSizable;
  tv.textContainer.widthTracksTextView = true;
  tv.textStorage.setAttributedString(out);
  sv.documentView = tv;

  var alert = $.NSAlert.alloc.init;
  alert.messageText = $('ViFormat');
  alert.addButtonWithTitle($('Đóng'));
  alert.accessoryView = sv;
  if (argv[1] === 'auto') {
    $.NSTimer.scheduledTimerWithTimeIntervalTargetSelectorUserInfoRepeats(1.5, $.NSApp, 'abortModal', null, false);
  }
  alert.runModal;
  return 'ok';
}
