/*
 * Company Signature Add-in — event-based runtime handler.
 *
 * Runs automatically when a new message compose window opens — for new
 * mail AND replies/forwards, since OnNewMessageCompose fires for both.
 * Uses getComposeTypeAsync() to tell them apart, then applies a
 * separate template (and separate opt-out) for each case.
 *
 * This file is also loaded directly (via <script src="functions.js">)
 * by preview.html on GitHub Pages, outside of Outlook entirely — so all
 * Office.js-specific calls below are guarded to only run when Office.js
 * is actually present.
 */

if (typeof Office !== "undefined") {
  /*
   * Office.js requires onReady (or Office.initialize) to be called as
   * part of the page's load sequence — without this, Office.js
   * considers itself not fully loaded and throws even if the rest of
   * the script is correct.
   */
  Office.onReady();
}

// ---------------------------------------------------------------------
// TEMPORARY TROUBLESHOOTING FLAG
// When true, a fetch failure (e.g. CORS, 404, network error) inserts a
// visible debug message into the signature instead of silently doing
// nothing — useful for confirming whether the directory.json fetch is
// actually the problem. Set back to false (or delete this whole block,
// see below) once troubleshooting is done.
// ---------------------------------------------------------------------
const DEBUG_FALLBACK_ENABLED = false;

function buildDebugFallbackHtml(errorMessage) {
  return (
    '<table cellpadding="0" cellspacing="0" border="0">' +
      '<tr><td style="font-family:Arial, sans-serif; font-size:12px; ' +
      'color:#B00020; border:1px dashed #B00020; padding:6px;">' +
      'DEBUG: signature directory fetch failed — ' + escapeHtml(errorMessage) + '. ' +
      'This message only appears because DEBUG_FALLBACK_ENABLED is true in functions.js.' +
      '</td></tr>' +
    '</table>'
  );
}
// ---------------------------------------------------------------------
// END TEMPORARY TROUBLESHOOTING BLOCK — delete the two items above
// (the flag and the function) along with the "if (DEBUG_FALLBACK_ENABLED)"
// check inside the .catch() below, once you're done troubleshooting.
// ---------------------------------------------------------------------

// ---------------------------------------------------------------------
// EDIT THIS: point at your hosted directory.json.
// ---------------------------------------------------------------------
const SIGNATURE_URL = "https://sig.resilift.com.au/";
const DIRECTORY_URL = SIGNATURE_URL + "directory.json?v=1";

// ---------------------------------------------------------------------
// Optional free-text/HTML sections around the signature — "nb" (notice)
// above, "f" (footer) below (e.g. "I only work Mondays and Fridays
// 9-3pm" as a notice, or a disclaimer/promo line as a footer). Both are
// completely left out (no wrapping div at all) if empty, missing, or
// whitespace-only. Rendered as raw HTML (not escaped) so you can put
// links, bold text, <br> line breaks, etc. in the sheet cell — this is
// safe because directory.json only ever comes from your own controlled
// sheet, not user-submitted input.
// ---------------------------------------------------------------------
function buildNoticeHtml(user) {
  if (!user.nb || user.nb.toString().trim() === "") return "";
  return (
    // '<div style="font-family: Arial, Helvetica, sans-serif; font-size: 13px; line-height: 1.3; color: rgb(0, 0, 0); padding-top: 16px;">' + user.nb + '</div>' + // doesn't work
    '<div style="margin-top: 16px; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 17px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">' + user.nb + '</div>'  // works
    // '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' // works
  );
}
// '<div style="margin-top: 16px; font-family: Arial, Helvetica, sans-serif; font-size: 13px; line-height: 17px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><i>Please note I only work Mon-Wed 9am-3pm</i></div>' +

function buildFooterHtml(user) {
  if (!user.f || user.f.toString().trim() === "") return "";
  return (
    // '<div style="font-family: Arial, Helvetica, sans-serif; font-size: 13px; ' +
    // 'line-height: 1.3; color: rgb(0, 0, 0); padding-bottom: 16px">' + user.f + '</div>'
    // '<div style="margin-bottom: 16px; font-family: Arial, Helvetica, sans-serif; font-size: 13px; line-height: 17px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">' + user.f + '</div>'  // works
    '<div>' + user.f + '</div>' + // works
    '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>'
  );
}

// ---------------------------------------------------------------------
// TEMPLATES — a single registry used for BOTH new mail ("te" field) and
// replies/forwards ("rte" field) in directory.json. Any key here can be
// used in either field — there's no enforcement that a "reply_"-
// prefixed template is only used for replies, that's just a naming
// convention if you want to keep track of which ones you intend for
// replies vs new mail.
//
// Directory field key reference (obfuscated short keys used in
// directory.json, mapped from the full names):
//   n   = name
//   t   = title
//   tl  = titleLine
//   l   = location
//   pd  = phoneDisplay
//   pl  = phoneLink
//   d   = distributor
//   te  = template
//   er  = excludeFromReplies
//   rte = replyTemplate
//   nb  = notice
//   f   = footer
// ---------------------------------------------------------------------
const TEMPLATES = {

  // hamish: function (user) {
  //   return (
  //     buildNoticeHtml(user) +
  //     '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' + // works
  //     '<table role="presentation" cellspacing="0" cellpadding="0" border="0" style="width:360px; max-width:360px; box-sizing:border-box; border-collapse:collapse; border-spacing:0px">' +
  //       '<tbody>' +
  //           '<tr>' +
  //               '<td style="vertical-align:middle; width:110px">' +
  //                   '<div style="font-family: Arial, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);"><img src="' + SIGNATURE_URL + 'assets/RESiLIFT_signature_logo_220px.png" alt="RESiLIFT Logo" width="110" height="81" style="width: 110px; height: 81px; display: block;"></div>' +
  //               '</td>' +
  //               '<td style="padding-left:10px; vertical-align:middle; width:250px">' +
  //                 '<table cellspacing="0" cellpadding="0" style="box-sizing:border-box; border-collapse:collapse; border-spacing:0px">' +
  //                   '<tbody><tr><td>' +
  //                     '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 21px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><b style="">' + escapeHtml(user.n) + '</b></div>' +
  //                     '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 20px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">' + escapeHtml(user.tl) + '</div>' +
  //                     '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 18px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="tel:' + escapeHtml(user.pl) + '" target="_blank" style="color: rgb(0, 0, 0); text-decoration: none;">' + escapeHtml(user.pd) + '</a></div>' +
  //                     '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 18px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="tel:1300303522" target="_blank" style="color: rgb(0, 0, 0); text-decoration: none;">1300 303 522</a></div>' +
  //                     '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="https://resilift.com.au/" target="_blank" style="color: rgb(0, 0, 0); text-decoration: none;">resilift.com.au</a></div>' +                    '</td></tr></tbody>' +
  //                 '</table>' +
  //               '</td>' +
  //           '</tr>' +
  //       '</tbody>' +
  //     '</table>' +
  //     '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' + // works
  //     buildFooterHtml(user) +
  //     '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' + // works
  //     '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' + // works
  //     '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' + // works

  //     buildNoticeHtml(user) +
  //     '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' +
  //     '<table role="presentation" cellspacing="0" cellpadding="0" border="0" ' +
  //     'style="width:360px; max-width:360px; box-sizing:border-box; border-collapse:collapse; border-spacing:0px">' +
  //       '<tbody><tr>' +
  //         '<td style="vertical-align:middle; width:120px">' +
  //           '<div style="font-family: Arial, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">' +
  //             '<img src="' + SIGNATURE_URL + 'assets/RESiLIFT_signature_logo_240px.png" ' +
  //             'alt="RESiLIFT Logo" width="120" height="88" style="width: 120px; height: 88px; display: block;">' +
  //           '</div>' +
  //         '</td>' +
  //         '<td style="padding-left:11px; vertical-align:middle; width:240px">' +
  //           '<table cellspacing="0" cellpadding="0" style="box-sizing:border-box; border-collapse:collapse; border-spacing:0px">' +
  //             '<tbody><tr><td>' +
  //               '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 21px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><b>' + escapeHtml(user.n) + '</b></div>' +
  //               '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 20px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">' + escapeHtml(user.tl) + '</div>' +
  //               '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 18px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">' + escapeHtml(user.d) + '</div>' +
  //               '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 19px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">Authorised RESiLIFT Distributor</div>' +
  //               '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 18px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="tel:' + escapeHtml(user.pl) + '" target="_blank" style="color: rgb(0, 0, 0); text-decoration: none;">' + escapeHtml(user.pd) + '</a></div>' +
  //               '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 18px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="tel:1300303522" target="_blank" style="color: rgb(0, 0, 0); text-decoration: none;">1300 303 522</a></div>' +
  //               '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="https://resilift.com.au/" style="color: rgb(0, 0, 0); text-decoration: none;">resilift.com.au</a></div>' +
  //             '</td></tr></tbody>' +
  //           '</table>' +
  //         '</td>' +
  //       '</tr></tbody>' +
  //     '</table>' +
  //     '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' +
  //     buildFooterHtml(user) +
  //     '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' + // works
  //     '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' + // works
  //     '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' + // works

  //     buildNoticeHtml(user) +
  //     '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' +
  //     '<table role="presentation" cellspacing="0" cellpadding="0" border="0" ' +
  //     'style="width:360px; max-width:360px; box-sizing:border-box; border-collapse:collapse; border-spacing:0px">' +
  //       '<tbody><tr>' +
  //         '<td style="vertical-align:middle; width:120px">' +
  //           '<div style="font-family: Arial, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">' +
  //             '<img src="' + SIGNATURE_URL + 'assets/RESiLIFT_Manufacturing_Logo.png" ' +
  //             'alt="RESiLIFT Logo" width="120" height="96"style="width: 120px; height: 95px; display: block;">' +
  //           '</div>' +
  //         '</td>' +
  //         '<td style="padding-left:10px; vertical-align:middle; width:240px">' +
  //           '<table cellspacing="0" cellpadding="0" style="box-sizing:border-box; border-collapse:collapse; border-spacing:0px">' +
  //             '<tbody><tr><td>' +
  //               '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 22px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><b>' + escapeHtml(user.n) + '</b></div>' +
  //               '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 21px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">' + escapeHtml(user.tl) + '</div>' +
  //               '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 20px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="tel:' + escapeHtml(user.pl) + '" target="_blank" style="color: rgb(0, 0, 0); text-decoration: none;">' + escapeHtml(user.pd) + '</a></div>' +
  //               '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 20px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="tel:1300303522" target="_blank" style="color: rgb(0, 0, 0); text-decoration: none;">1300 303 522</a></div>' +
  //               '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 17px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="https://resilift.com.au/" style="color: rgb(0, 0, 0); text-decoration: none;">resilift.com.au</a>' +
  //               '</div>' +
  //             '</td></tr></tbody>' +
  //           '</table>' +
  //         '</td>' +
  //       '</tr></tbody>' +
  //     '</table>' +
  //     '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' +
  //     buildFooterHtml(user)
  //   );
  // },

  simple: function (user) {
    return (
      buildNoticeHtml(user) +
      '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' + // works
      '<table role="presentation" cellspacing="0" cellpadding="0" border="0" style="width:360px; max-width:360px; box-sizing:border-box; border-collapse:collapse; border-spacing:0px">' +
        '<tbody>' +
            '<tr>' +
                '<td style="vertical-align:middle; width:110px">' +
                    '<div style="font-family: Arial, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);"><img src="' + SIGNATURE_URL + 'assets/RESiLIFT_signature_logo_220px.png" alt="RESiLIFT Logo" width="110" height="81" style="width: 110px; height: 81px; display: block;"></div>' +
                '</td>' +
                '<td style="padding-left:10px; vertical-align:middle; width:250px">' +
                  '<table cellspacing="0" cellpadding="0" style="box-sizing:border-box; border-collapse:collapse; border-spacing:0px">' +
                    '<tbody><tr><td>' +
                      '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 21px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><b style="">' + escapeHtml(user.n) + '</b></div>' +
                      '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 20px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">' + escapeHtml(user.tl) + '</div>' +
                      '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 18px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="tel:' + escapeHtml(user.pl) + '" target="_blank" style="color: rgb(0, 0, 0); text-decoration: none;">' + escapeHtml(user.pd) + '</a></div>' +
                      '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 18px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="tel:1300303522" target="_blank" style="color: rgb(0, 0, 0); text-decoration: none;">1300 303 522</a></div>' +
                      '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="https://resilift.com.au/" target="_blank" style="color: rgb(0, 0, 0); text-decoration: none;">resilift.com.au</a></div>' +                    '</td></tr></tbody>' +
                  '</table>' +
                '</td>' +
            '</tr>' +
        '</tbody>' +
      '</table>' +
      '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' + // works
      buildFooterHtml(user)
    );
  },

  distributor: function (user) {
    return (
      buildNoticeHtml(user) +
      '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' +
      '<table role="presentation" cellspacing="0" cellpadding="0" border="0" ' +
      'style="width:360px; max-width:360px; box-sizing:border-box; border-collapse:collapse; border-spacing:0px">' +
        '<tbody><tr>' +
          '<td style="vertical-align:middle; width:120px">' +
            '<div style="font-family: Arial, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">' +
              '<img src="' + SIGNATURE_URL + 'assets/RESiLIFT_signature_logo_240px.png" ' +
              'alt="RESiLIFT Logo" width="120" height="88" style="width: 120px; height: 88px; display: block;">' +
            '</div>' +
          '</td>' +
          '<td style="padding-left:11px; vertical-align:middle; width:240px">' +
            '<table cellspacing="0" cellpadding="0" style="box-sizing:border-box; border-collapse:collapse; border-spacing:0px">' +
              '<tbody><tr><td>' +
                '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 21px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><b>' + escapeHtml(user.n) + '</b></div>' +
                '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 20px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">' + escapeHtml(user.tl) + '</div>' +
                '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 18px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">' + escapeHtml(user.d) + '</div>' +
                '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 19px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">Authorised RESiLIFT Distributor</div>' +
                '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 18px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="tel:' + escapeHtml(user.pl) + '" target="_blank" style="color: rgb(0, 0, 0); text-decoration: none;">' + escapeHtml(user.pd) + '</a></div>' +
                '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 18px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="tel:1300303522" target="_blank" style="color: rgb(0, 0, 0); text-decoration: none;">1300 303 522</a></div>' +
                '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="https://resilift.com.au/" style="color: rgb(0, 0, 0); text-decoration: none;">resilift.com.au</a></div>' +
              '</td></tr></tbody>' +
            '</table>' +
          '</td>' +
        '</tr></tbody>' +
      '</table>' +
      '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' +
      buildFooterHtml(user)
    );
  },

  manufacturer: function (user) {
    return (
      buildNoticeHtml(user) +
      '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' +
      '<table role="presentation" cellspacing="0" cellpadding="0" border="0" ' +
      'style="width:360px; max-width:360px; box-sizing:border-box; border-collapse:collapse; border-spacing:0px">' +
        '<tbody><tr>' +
          '<td style="vertical-align:middle; width:120px">' +
            '<div style="font-family: Arial, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">' +
              '<img src="' + SIGNATURE_URL + 'assets/RESiLIFT_Manufacturing_Logo.png" ' +
              'alt="RESiLIFT Logo" width="120" height="96"style="width: 120px; height: 95px; display: block;">' +
            '</div>' +
          '</td>' +
          '<td style="padding-left:10px; vertical-align:middle; width:240px">' +
            '<table cellspacing="0" cellpadding="0" style="box-sizing:border-box; border-collapse:collapse; border-spacing:0px">' +
              '<tbody><tr><td>' +
                '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 22px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><b>' + escapeHtml(user.n) + '</b></div>' +
                '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 21px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">' + escapeHtml(user.tl) + '</div>' +
                '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 20px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="tel:' + escapeHtml(user.pl) + '" target="_blank" style="color: rgb(0, 0, 0); text-decoration: none;">' + escapeHtml(user.pd) + '</a></div>' +
                '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 20px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="tel:1300303522" target="_blank" style="color: rgb(0, 0, 0); text-decoration: none;">1300 303 522</a></div>' +
                '<div style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 17px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);"><a href="https://resilift.com.au/" style="color: rgb(0, 0, 0); text-decoration: none;">resilift.com.au</a>' +
                '</div>' +
              '</td></tr></tbody>' +
            '</table>' +
          '</td>' +
        '</tr></tbody>' +
      '</table>' +
      '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' +
      buildFooterHtml(user)
    );
  },

  test: function(user) {
    const u2 = { ...user, tl: user.t };

    return (
      TEMPLATES.simple(user) +
      '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' +
      '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' +
      TEMPLATES.distributor(u2) +
      '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' +
      '<p style="margin: 0; font-size: 16px; line-height: 16px; mso-line-height-rule: exactly; color: rgb(0, 0, 0);">&nbsp;</p>' +
      TEMPLATES.manufacturer(u2)
    );
  }

};

const DEFAULT_TEMPLATE_KEY = "simple";

function buildSignatureHtml(user) {
  const renderer = TEMPLATES[user.te] || TEMPLATES[DEFAULT_TEMPLATE_KEY];
  return renderer(user);
}

function buildReplySignatureHtml(user) {
  const renderer = TEMPLATES[user.rte] || TEMPLATES[DEFAULT_TEMPLATE_KEY];
  return renderer(user);
}

function escapeHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Expose the pieces preview.html (or any other page) needs, so that
// page can just <script src="functions.js"> and reuse the exact same
// templates rather than duplicating them.
if (typeof window !== "undefined") {
  window.SignatureAddin = {
    TEMPLATES: TEMPLATES,
    DEFAULT_TEMPLATE_KEY: DEFAULT_TEMPLATE_KEY,
    DIRECTORY_URL: DIRECTORY_URL,
    buildSignatureHtml: buildSignatureHtml,
    buildReplySignatureHtml: buildReplySignatureHtml,
    escapeHtml: escapeHtml
  };
}

// ---------------------------------------------------------------------
// Event handler — fires on new message compose AND on reply/forward
// (OnNewMessageCompose covers both; getComposeTypeAsync tells them apart).
// Only defined/used when actually running inside Outlook.
// ---------------------------------------------------------------------
function onNewMessageComposeHandler(event) {
  const item = Office.context.mailbox.item;
  const email = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();

  item.getComposeTypeAsync(function (composeTypeResult) {
    const isReplyOrForward =
      composeTypeResult.status === Office.AsyncResultStatus.Succeeded &&
      (composeTypeResult.value.composeType === Office.MailboxEnums.ComposeType.Reply ||
       composeTypeResult.value.composeType === Office.MailboxEnums.ComposeType.Forward);

    fetch(DIRECTORY_URL)
      .then(function (response) {
        if (!response.ok) throw new Error("Directory fetch failed: " + response.status);
        return response.json();
      })
      .then(function (directory) {
        const user = directory[email];

        // No matching entry in the directory at all — do nothing. Don't
        // call setSignatureAsync, don't touch the compose body. Whatever
        // signature Outlook itself already inserted (the user's own
        // configured signature, if any) is left exactly as it was.
        if (!user) {
          event.completed();
          return;
        }

        if (isReplyOrForward) {
          // "er" (excludeFromReplies) defaults to false if not explicitly set.
          const excludeFromReplies = user.er === true;

          if (excludeFromReplies || user.rte === null || user.rte === "none") {
            // Do nothing — leave any existing signature on the reply as-is.
            event.completed();
            return;
          }

          insertSignature(buildReplySignatureHtml(user), event);
          return;
        }

        // New mail.
        if (user.te === null || user.te === "none") {
          event.completed();
          return;
        }

        insertSignature(buildSignatureHtml(user), event);
      })
      .catch(function (err) {
        // Directory fetch itself failed (network issue, CORS, 404, etc.).
        console.error("Signature directory lookup failed:", err);

        // TEMPORARY: see DEBUG_FALLBACK_ENABLED block above — remove this
        // if-branch (keep just the else branch's event.completed()) once
        // troubleshooting is done.
        if (DEBUG_FALLBACK_ENABLED) {
          insertSignature(buildDebugFallbackHtml(err.message || String(err)), event);
        } else {
          event.completed();
        }
      });
  });
}

function insertSignature(html, event) {
  Office.context.mailbox.item.body.setSignatureAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(
          "setSignatureAsync failed: " +
          asyncResult.error.code + " " + asyncResult.error.message
        );
      }
      event.completed();
    }
  );
}

if (typeof Office !== "undefined") {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}
