// Mobile menu toggle
const menuToggle = document.querySelector('.menu-toggle');
const siteNav = document.querySelector('.site-nav');
if (menuToggle && siteNav) {
  menuToggle.addEventListener('click', () => {
    const open = siteNav.classList.toggle('open');
    menuToggle.setAttribute('aria-expanded', String(open));
  });
  siteNav.querySelectorAll('a').forEach(a => a.addEventListener('click', () => {
    siteNav.classList.remove('open');
    menuToggle.setAttribute('aria-expanded', 'false');
  }));
}

// Reading progress bar
const progressBar = document.getElementById('progressBar');
const updateProgress = () => {
  if (!progressBar) return;
  const max = document.documentElement.scrollHeight - window.innerHeight;
  const value = max > 0 ? (window.scrollY / max) * 100 : 0;
  progressBar.style.width = `${value}%`;
};
window.addEventListener('scroll', updateProgress, { passive: true });
updateProgress();

// Toast Notification
const toast = document.getElementById('toast');
let toastTimer;
function showToast(message = 'Copied to Clipboard!') {
  if (!toast) return;
  toast.textContent = message;
  toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove('show'), 2200);
}

// Copy Buttons
document.querySelectorAll('[data-copy]').forEach(button => {
  button.addEventListener('click', async () => {
    const text = button.getAttribute('data-copy') || '';
    try {
      await navigator.clipboard.writeText(text);
      showToast('Copied to Clipboard! ✨');
    } catch {
      const temp = document.createElement('textarea');
      temp.value = text;
      document.body.appendChild(temp);
      temp.select();
      document.execCommand('copy');
      temp.remove();
      showToast('Copied to Clipboard! ✨');
    }
  });
});

// 3-Visual Image Tab Switcher & Dynamic Download URL Updater
document.querySelectorAll('.visual-tabs-bar').forEach(bar => {
  const tabs = bar.querySelectorAll('.img-tab-btn');
  const storyVisualContainer = bar.closest('.story-visual-phone');
  const phoneImage = storyVisualContainer ? storyVisualContainer.querySelector('.phone-mockup img') : null;
  const downloadBtn = storyVisualContainer ? storyVisualContainer.querySelector('.download-img-btn') : null;

  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      tabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');

      const newSrc = tab.getAttribute('data-img-src');
      if (phoneImage && newSrc) {
        phoneImage.style.opacity = '0.3';
        setTimeout(() => {
          phoneImage.src = newSrc;
          phoneImage.style.opacity = '1';
        }, 150);
      }

      // Update download link
      if (downloadBtn && newSrc) {
        downloadBtn.setAttribute('href', newSrc);
        const filename = newSrc.split('/').pop();
        downloadBtn.setAttribute('download', filename);
      }
    });
  });
});

// Category Filter Bar Logic
const filterButtons = document.querySelectorAll('.filter-btn');
const storyCards = document.querySelectorAll('.story-card');
const stories = document.querySelectorAll('.story');

filterButtons.forEach(btn => {
  btn.addEventListener('click', () => {
    filterButtons.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    const category = btn.getAttribute('data-category');

    storyCards.forEach(card => {
      const cardCat = card.getAttribute('data-category');
      if (category === 'all' || cardCat === category) {
        card.style.display = 'flex';
      } else {
        card.style.display = 'none';
      }
    });

    stories.forEach(story => {
      const storyCat = story.getAttribute('data-category');
      if (category === 'all' || storyCat === category) {
        story.style.display = 'grid';
      } else {
        story.style.display = 'none';
      }
    });
  });
});

// Teleprompter Script Modal
const scriptModal = document.getElementById('scriptModal');
const scriptModalTitle = document.getElementById('scriptModalTitle');
const scriptModalContent = document.getElementById('scriptModalContent');
const closeModalBtn = document.getElementById('closeScriptModal');

document.querySelectorAll('.open-teleprompter-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    const storyElem = btn.closest('.story');
    if (!storyElem) return;
    
    const storyTitle = storyElem.querySelector('h2')?.textContent || 'Reel Video Script';
    const scriptTableRows = storyElem.querySelectorAll('tbody tr');

    if (scriptModalTitle) scriptModalTitle.textContent = `📱 Teleprompter: ${storyTitle}`;
    
    if (scriptModalContent) {
      scriptModalContent.innerHTML = '';
      scriptTableRows.forEach(row => {
        const cells = row.querySelectorAll('td');
        if (cells.length >= 3) {
          const timing = cells[0].textContent;
          const visual = cells[1].textContent;
          const narration = cells[2].textContent;

          const lineDiv = document.createElement('div');
          lineDiv.className = 'prompter-line';
          lineDiv.innerHTML = `
            <span class="time">⏱️ ${timing}</span>
            <span class="vis">🎬 Visual: ${visual}</span>
            <span class="narr">🗣️ "${narration}"</span>
          `;
          scriptModalContent.appendChild(lineDiv);
        }
      });
    }

    if (scriptModal) {
      scriptModal.classList.add('open');
      scriptModal.setAttribute('aria-hidden', 'false');
      document.body.style.overflow = 'hidden';
    }
  });
});

if (closeModalBtn && scriptModal) {
  closeModalBtn.addEventListener('click', () => {
    scriptModal.classList.remove('open');
    scriptModal.setAttribute('aria-hidden', 'true');
    document.body.style.overflow = '';
  });
  scriptModal.addEventListener('click', (e) => {
    if (e.target === scriptModal) {
      scriptModal.classList.remove('open');
      scriptModal.setAttribute('aria-hidden', 'true');
      document.body.style.overflow = '';
    }
  });
}
