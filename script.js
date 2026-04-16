const exportPdfBtn = document.getElementById("exportPdfBtn");
const exportPptxBtn = document.getElementById("exportPptxBtn");
const navLinks = Array.from(document.querySelectorAll(".quick-nav__link"));
const menuToggle = document.getElementById("menuToggle");
const quickNav = document.getElementById("primary-navigation");
const navBackdrop = document.getElementById("navBackdrop");
const sections = Array.from(document.querySelectorAll(".panel[id]"));

if (menuToggle && quickNav) {
  menuToggle.addEventListener("click", () => {
    const isOpen = menuToggle.getAttribute("aria-expanded") === "true";
    const nextOpen = !isOpen;
    menuToggle.setAttribute("aria-expanded", String(nextOpen));
    menuToggle.classList.toggle("is-open", nextOpen);
    quickNav.classList.toggle("is-open", nextOpen);
    navBackdrop?.classList.toggle("is-visible", nextOpen);
  });
}

function closeMobileMenu() {
  if (!menuToggle || !quickNav) return;
  menuToggle.setAttribute("aria-expanded", "false");
  menuToggle.classList.remove("is-open");
  quickNav.classList.remove("is-open");
  navBackdrop?.classList.remove("is-visible");
}

if (navBackdrop) {
  navBackdrop.addEventListener("click", closeMobileMenu);
}

function setActiveLink(id) {
  navLinks.forEach((link) => {
    const isMatch = link.getAttribute("href") === `#${id}`;
    link.classList.toggle("is-active", isMatch);
  });
}

const sectionObserver = new IntersectionObserver(
  (entries) => {
    entries.forEach((entry) => {
      if (entry.isIntersecting) {
        setActiveLink(entry.target.id);
      }
    });
  },
  {
    root: null,
    threshold: 0.35,
    rootMargin: "-10% 0px -45% 0px"
  }
);

sections.forEach((section) => sectionObserver.observe(section));

navLinks.forEach((link) => {
  link.addEventListener("click", closeMobileMenu);
});

if (exportPdfBtn) {
  exportPdfBtn.addEventListener("click", () => {
    window.print();
  });
}

function setExportBusyState(button, isBusy, busyText, idleText) {
  if (!button) return;
  button.disabled = isBusy;
  button.textContent = isBusy ? busyText : idleText;
}

async function renderSectionToImage(section) {
  const canvas = await window.html2canvas(section, {
    scale: 2,
    useCORS: true,
    backgroundColor: "#ffffff",
    scrollX: 0,
    scrollY: -window.scrollY
  });
  return canvas.toDataURL("image/png");
}

async function exportToPptx() {
  if (!window.html2canvas || !window.PptxGenJS) {
    alert("تعذر تحميل أدوات التصدير. تأكد من الاتصال بالإنترنت ثم أعد المحاولة.");
    return;
  }

  const pptx = new window.PptxGenJS();
  pptx.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 (16:9)
  pptx.author = "Yemen Soft";
  pptx.subject = "عرض استخدام أدوات الذكاء الاصطناعي في الدعم الفني";
  pptx.title = "استخدام أدوات الذكاء الاصطناعي في الدعم الفني";
  pptx.lang = "ar-SA";

  setExportBusyState(exportPptxBtn, true, "جاري تجهيز PowerPoint...", "تصدير إلى PowerPoint");

  try {
    const slideW = 13.333;
    const slideH = 7.5;

    for (const section of sections) {
      const slide = pptx.addSlide();
      const imageData = await renderSectionToImage(section);
      const sectionWidth = section.offsetWidth || 1;
      const sectionHeight = section.offsetHeight || 1;
      const sectionRatio = sectionWidth / sectionHeight;
      const slideRatio = slideW / slideH;

      let drawW = slideW;
      let drawH = slideH;
      let offsetX = 0;
      let offsetY = 0;

      // عرض الصورة داخل الشريحة دون أي تمدد أو تشويه.
      if (sectionRatio > slideRatio) {
        drawH = slideW / sectionRatio;
        offsetY = (slideH - drawH) / 2;
      } else {
        drawW = slideH * sectionRatio;
        offsetX = (slideW - drawW) / 2;
      }

      slide.background = { color: "FFFFFF" };
      slide.addImage({
        data: imageData,
        x: offsetX,
        y: offsetY,
        w: drawW,
        h: drawH
      });
    }

    await pptx.writeFile({ fileName: "عرض-تنفيذي-استخدام-ادوات-AI.pptx" });
  } catch (error) {
    console.error(error);
    alert("حدث خطأ أثناء إنشاء ملف PowerPoint. حاول مرة أخرى.");
  } finally {
    setExportBusyState(exportPptxBtn, false, "جاري تجهيز PowerPoint...", "تصدير إلى PowerPoint");
  }
}

if (exportPptxBtn) {
  exportPptxBtn.addEventListener("click", exportToPptx);
}
