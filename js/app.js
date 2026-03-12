(function () {
  'use strict';

  // --- State ---
  let names = [];
  let spinning = false;
  let currentAngle = 0;

  // --- Color palette ---
  const COLORS = [
    '#667eea', '#764ba2', '#f093fb', '#f5576c',
    '#4facfe', '#00f2fe', '#43e97b', '#fa709a',
    '#fee140', '#ff9a9e', '#a18cd1', '#fbc2eb',
    '#ffecd2', '#fcb69f', '#a1c4fd', '#c2e9fb',
    '#d4fc79', '#96e6a1', '#dfe6e9', '#fda085',
  ];

  // --- DOM refs ---
  const canvas = document.getElementById('wheelCanvas');
  const ctx = canvas.getContext('2d');
  const spinBtn = document.getElementById('spinBtn');
  const namesInput = document.getElementById('namesInput');
  const updateBtn = document.getElementById('updateBtn');
  const shuffleBtn = document.getElementById('shuffleBtn');
  const clearBtn = document.getElementById('clearBtn');
  const winnerOverlay = document.getElementById('winnerOverlay');
  const winnerNameEl = document.getElementById('winnerName');
  const closeWinnerBtn = document.getElementById('closeWinner');
  const confettiContainer = document.getElementById('confetti');

  // --- Teams SDK init ---
  function initTeams() {
    try {
      if (window.microsoftTeams) {
        microsoftTeams.app.initialize().then(function () {
          microsoftTeams.app.getContext().then(function (context) {
            if (context.app.theme === 'dark') {
              document.body.classList.add('theme-dark');
            }

            // If loaded as a configurable tab config page, enable the Save button
            if (context.page.frameContext === 'settings') {
              microsoftTeams.pages.config.registerOnSaveHandler(function (saveEvent) {
                microsoftTeams.pages.config.setConfig({
                  suggestedDisplayName: 'Will bot',
                  entityId: 'wheelOfNames',
                  contentUrl: window.location.origin + window.location.pathname,
                  websiteUrl: window.location.origin + window.location.pathname
                });
                saveEvent.notifySuccess();
              });
              microsoftTeams.pages.config.setValidityState(true);
            }
          });
        });
      }
    } catch (e) {
      // Running outside Teams — that's fine
    }
  }

  // --- Wheel Drawing ---
  function drawWheel() {
    var size = canvas.width;
    var center = size / 2;
    var radius = center - 8;

    ctx.clearRect(0, 0, size, size);

    if (names.length === 0) {
      drawEmptyWheel(center, radius);
      return;
    }

    var sliceAngle = (2 * Math.PI) / names.length;

    names.forEach(function (name, i) {
      var startAngle = currentAngle + i * sliceAngle;
      var endAngle = startAngle + sliceAngle;

      // Draw segment
      ctx.beginPath();
      ctx.moveTo(center, center);
      ctx.arc(center, center, radius, startAngle, endAngle);
      ctx.closePath();
      ctx.fillStyle = COLORS[i % COLORS.length];
      ctx.fill();

      // Segment border
      ctx.strokeStyle = 'rgba(0, 0, 0, 0.2)';
      ctx.lineWidth = 2;
      ctx.stroke();

      // Draw name text
      ctx.save();
      ctx.translate(center, center);
      ctx.rotate(startAngle + sliceAngle / 2);

      ctx.textAlign = 'right';
      ctx.fillStyle = getContrastColor(COLORS[i % COLORS.length]);
      ctx.font = getFontSize(names.length) + 'px "Segoe UI", sans-serif';
      ctx.fillText(truncateName(name, names.length), radius - 16, 5);

      ctx.restore();
    });

    // Center circle
    ctx.beginPath();
    ctx.arc(center, center, 24, 0, 2 * Math.PI);
    ctx.fillStyle = '#1a1a2e';
    ctx.fill();
    ctx.strokeStyle = '#667eea';
    ctx.lineWidth = 3;
    ctx.stroke();
  }

  function drawEmptyWheel(center, radius) {
    ctx.beginPath();
    ctx.arc(center, center, radius, 0, 2 * Math.PI);
    ctx.fillStyle = '#16213e';
    ctx.fill();
    ctx.strokeStyle = '#2a2a4a';
    ctx.lineWidth = 3;
    ctx.stroke();

    ctx.fillStyle = '#5a5a7a';
    ctx.font = '18px "Segoe UI", sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Add team members and', center, center - 12);
    ctx.fillText('click Update Wheel', center, center + 16);
    ctx.fillStyle = '#4a4a6a';
    ctx.font = '14px "Segoe UI", sans-serif';
    ctx.fillText('Who tells the next story?', center, center + 44);
  }

  function getFontSize(count) {
    if (count <= 6) return 16;
    if (count <= 12) return 13;
    if (count <= 20) return 11;
    return 9;
  }

  function truncateName(name, count) {
    var maxLen = count <= 8 ? 18 : count <= 16 ? 12 : 8;
    return name.length > maxLen ? name.substring(0, maxLen - 1) + '\u2026' : name;
  }

  function getContrastColor(hex) {
    var r = parseInt(hex.slice(1, 3), 16);
    var g = parseInt(hex.slice(3, 5), 16);
    var b = parseInt(hex.slice(5, 7), 16);
    var luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    return luminance > 0.6 ? '#1a1a2e' : '#ffffff';
  }

  // --- Spin Logic ---
  function spin() {
    if (spinning || names.length < 2) return;

    spinning = true;
    spinBtn.disabled = true;

    // Random total rotation: 5-10 full spins + random offset
    var extraSpins = 5 + Math.random() * 5;
    var targetAngle = currentAngle + extraSpins * 2 * Math.PI + Math.random() * 2 * Math.PI;
    var startAngle = currentAngle;
    var totalDelta = targetAngle - startAngle;
    var duration = 4000 + Math.random() * 2000; // 4-6 seconds
    var startTime = null;

    function easeOutCubic(t) {
      return 1 - Math.pow(1 - t, 3);
    }

    function animate(timestamp) {
      if (!startTime) startTime = timestamp;
      var elapsed = timestamp - startTime;
      var progress = Math.min(elapsed / duration, 1);
      var easedProgress = easeOutCubic(progress);

      currentAngle = startAngle + totalDelta * easedProgress;
      drawWheel();

      if (progress < 1) {
        requestAnimationFrame(animate);
      } else {
        spinning = false;
        spinBtn.disabled = false;
        currentAngle = currentAngle % (2 * Math.PI);
        announceWinner();
      }
    }

    requestAnimationFrame(animate);
  }

  function announceWinner() {
    // The pointer is at the top, which is -PI/2 in canvas coordinates.
    // Segment i is drawn from (currentAngle + i * sliceAngle) to (currentAngle + (i+1) * sliceAngle).
    // We need to find which segment contains the angle -PI/2.
    var sliceAngle = (2 * Math.PI) / names.length;
    // Normalize (-PI/2 - currentAngle) into [0, 2PI)
    var pointerAngle = ((-Math.PI / 2 - currentAngle) % (2 * Math.PI) + 2 * Math.PI) % (2 * Math.PI);

    var winnerIndex = Math.floor(pointerAngle / sliceAngle) % names.length;

    winnerNameEl.textContent = names[winnerIndex];
    winnerOverlay.classList.remove('hidden');
    spawnConfetti();
  }

  // --- Confetti ---
  function spawnConfetti() {
    confettiContainer.innerHTML = '';
    var confettiColors = ['#667eea', '#764ba2', '#f5576c', '#43e97b', '#fee140', '#4facfe', '#fa709a'];

    for (var i = 0; i < 50; i++) {
      var piece = document.createElement('div');
      piece.className = 'confetti-piece';
      piece.style.left = Math.random() * 100 + '%';
      piece.style.backgroundColor = confettiColors[Math.floor(Math.random() * confettiColors.length)];
      piece.style.animationDelay = Math.random() * 2 + 's';
      piece.style.animationDuration = 2 + Math.random() * 2 + 's';
      piece.style.borderRadius = Math.random() > 0.5 ? '50%' : '0';
      piece.style.width = 6 + Math.random() * 8 + 'px';
      piece.style.height = 6 + Math.random() * 8 + 'px';
      confettiContainer.appendChild(piece);
    }
  }

  // --- Name Management ---
  function updateNames() {
    var text = namesInput.value.trim();
    names = text
      .split('\n')
      .map(function (n) { return n.trim(); })
      .filter(function (n) { return n.length > 0; });
    drawWheel();
  }

  function shuffleNames() {
    for (var i = names.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var temp = names[i];
      names[i] = names[j];
      names[j] = temp;
    }
    namesInput.value = names.join('\n');
    drawWheel();
  }

  function clearNames() {
    names = [];
    namesInput.value = '';
    drawWheel();
  }

  // --- Event Listeners ---
  spinBtn.addEventListener('click', spin);
  updateBtn.addEventListener('click', updateNames);
  shuffleBtn.addEventListener('click', shuffleNames);
  clearBtn.addEventListener('click', clearNames);
  closeWinnerBtn.addEventListener('click', function () {
    winnerOverlay.classList.add('hidden');
  });

  // Close overlay on background click
  winnerOverlay.addEventListener('click', function (e) {
    if (e.target === winnerOverlay) {
      winnerOverlay.classList.add('hidden');
    }
  });

  // --- Init ---
  initTeams();
  drawWheel();
})();
