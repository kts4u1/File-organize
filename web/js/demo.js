/* Example photos for the demo, drawn on a canvas: a tool back with a QR
 * sticker, and one with a handwritten sticky note. Both carry a calibration
 * label dated two days ago. Engineers and serials are made up. */
(function () {
  "use strict";
  const TC = (window.TC = window.TC || {});

  function body(ctx, w, h) {
    const g = ctx.createLinearGradient(0, 0, w, h);
    g.addColorStop(0, "#e9b631");
    g.addColorStop(1, "#c98f1c");
    ctx.fillStyle = g;
    ctx.fillRect(0, 0, w, h);
    ctx.fillStyle = "rgba(150, 90, 20, 0.35)";
    ctx.font = "700 46px sans-serif";
    ctx.fillText("TO AVOID ELECTRICAL SHOCK, REMOVE", 60, h - 360);
    ctx.fillText("TEST LEADS BEFORE OPENING CASE", 90, h - 300);
    ctx.fillStyle = "#1b1b1b";
    ctx.beginPath();
    ctx.arc(w / 2, h - 140, 95, 0, Math.PI * 2);
    ctx.fill();
  }

  function label(ctx, x, y, angle, dates, cert) {
    const w = 560, hgt = 380;
    ctx.save();
    ctx.translate(x + w / 2, y + hgt / 2);
    ctx.rotate((angle * Math.PI) / 180);
    ctx.translate(-w / 2, -hgt / 2);
    ctx.fillStyle = "#f4f4f2";
    ctx.fillRect(0, 0, w, hgt);
    ctx.strokeStyle = "#222";
    ctx.lineWidth = 3;
    ctx.strokeRect(14, 14, w - 28, hgt - 28);
    ctx.fillStyle = "#222";
    ctx.font = "700 34px 'IBM Plex Sans KR', sans-serif";
    ctx.fillText("교 정 필 증", 200, 62);
    ctx.font = "italic 24px serif";
    ctx.fillText("Calibration Label", 190, 92);
    const rows = [["성적서번호", cert], ["교정일자", dates[0].replace(/-/g, ".")], ["차기교정예정일자", dates[1].replace(/-/g, ".")]];
    rows.forEach(([k, v], i) => {
      const ry = 110 + i * 70;
      ctx.beginPath();
      ctx.moveTo(14, ry);
      ctx.lineTo(w - 14, ry);
      ctx.moveTo(210, ry);
      ctx.lineTo(210, ry + 70);
      ctx.stroke();
      ctx.font = "500 22px 'IBM Plex Sans KR', sans-serif";
      ctx.fillText(k, 26, ry + 44);
      ctx.font = "600 38px 'IBM Plex Mono', monospace";
      ctx.fillText(v, 240, ry + 50);
    });
    ctx.beginPath();
    ctx.moveTo(14, 320);
    ctx.lineTo(w - 14, 320);
    ctx.stroke();
    ctx.font = "500 24px 'IBM Plex Sans KR', sans-serif";
    ctx.fillText("예시 교정기관 (주)샘플교정", 110, 356);
    ctx.restore();
  }

  function toFile(canvas, name) {
    return new Promise((res) => canvas.toBlob((b) => res(new File([b], name, { type: "image/jpeg" })), "image/jpeg", 0.9));
  }

  async function demoPhotos(records) {
    if (document.fonts) {
      await Promise.allSettled([
        document.fonts.load("64px 'Nanum Pen Script'", "이영희0123456789"),
        document.fonts.load("600 38px 'IBM Plex Mono'", "0123"),
        document.fonts.load("700 34px 'IBM Plex Sans KR'", "교정필증"),
      ]);
    }
    const calDate = new Date(Date.parse(TC.xlsx.todayIso()) - 2 * 86400000).toISOString().slice(0, 10);
    const qrRec = records.find((r) => r.serial === "31550075") || records[0];
    const handRec = records.find((r) => r.serial === "31550069") || records[1];
    const files = [];

    // 1) QR sticker + label
    {
      const c = TC.vision.makeCanvas(1200, 1600);
      const ctx = c.getContext("2d");
      body(ctx, c.width, c.height);
      ctx.save();
      ctx.translate(250, 170);
      ctx.rotate(-0.03);
      TC.stickers.drawSticker(ctx, qrRec, 0, 0, 14);
      ctx.restore();
      label(ctx, 300, 640, -2.5, [calDate, TC.xlsx.addMonths(calDate, qrRec.intervalMonths || 12)], "DA26-1041-007");
      files.push(await toFile(c, "예시_QR스티커.jpg"));
    }
    // 2) handwritten sticky note + label
    {
      const c = TC.vision.makeCanvas(1200, 1600);
      const ctx = c.getContext("2d");
      body(ctx, c.width, c.height);
      ctx.save();
      ctx.translate(200, 120);
      ctx.rotate(0.02);
      ctx.fillStyle = "#bcd3ee";
      ctx.fillRect(0, 0, 760, 470);
      ctx.fillStyle = "#23263a";
      ctx.font = "150px 'Nanum Pen Script', cursive";
      ctx.fillText(handRec.names[0] || "", 80, 190);
      ctx.font = "140px 'Nanum Pen Script', cursive";
      ctx.fillText(handRec.serial, 70, 380);
      ctx.restore();
      label(ctx, 320, 700, 1.5, [calDate, TC.xlsx.addMonths(calDate, handRec.intervalMonths || 12)], "DA26-1041-008");
      files.push(await toFile(c, "예시_손글씨메모.jpg"));
    }
    return files;
  }

  TC.demoPhotos = demoPhotos;
})();
