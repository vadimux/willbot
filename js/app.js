(function () {
  'use strict';

  // --- Config ---
  var CLIENT_ID = 'b41bc6ad-2fef-41bc-abea-196732e74ed1';
  var TENANT_ID = 'b41b72d0-4e9f-4c26-8a69-f949f367c91d';

  var teamsContext = null;
  var isInTeams = false;
  var accessToken = null; // stored after successful auth

  // --- State ---
  var names = [];
  var spinning = false;
  var currentAngle = 0;

  // --- Color palette ---
  var COLORS = [
    '#667eea', '#764ba2', '#f093fb', '#f5576c',
    '#4facfe', '#00f2fe', '#43e97b', '#fa709a',
    '#fee140', '#ff9a9e', '#a18cd1', '#fbc2eb',
    '#ffecd2', '#fcb69f', '#a1c4fd', '#c2e9fb',
    '#d4fc79', '#96e6a1', '#dfe6e9', '#fda085',
  ];

  // --- DOM refs ---
  var canvas = document.getElementById('wheelCanvas');
  var ctx = canvas.getContext('2d');
  var spinBtn = document.getElementById('spinBtn');
  var namesInput = document.getElementById('namesInput');
  var updateBtn = document.getElementById('updateBtn');
  var shuffleBtn = document.getElementById('shuffleBtn');
  var clearBtn = document.getElementById('clearBtn');
  var winnerOverlay = document.getElementById('winnerOverlay');
  var winnerNameEl = document.getElementById('winnerName');
  var closeWinnerBtn = document.getElementById('closeWinner');
  var confettiContainer = document.getElementById('confetti');
  var loadChatBtn = document.getElementById('loadChatBtn');
  var loadStatus = document.getElementById('loadStatus');
  var postToChatCheckbox = document.getElementById('postToChat');
  var postStatus = document.getElementById('postStatus');

  // --- Teams SDK init ---
  function initTeams() {
    try {
      if (window.microsoftTeams) {
        microsoftTeams.app.initialize().then(function () {
          isInTeams = true;
          microsoftTeams.app.getContext().then(function (context) {
            teamsContext = context;

            if (context.app.theme === 'dark') {
              document.body.classList.add('theme-dark');
            }

            // If loaded as a configurable tab config page, enable the Save button
            // Try silent auth on load
            trySilentAuth();

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
        }).catch(function () {
          isInTeams = false;
        });
      }
    } catch (e) {
      // Running outside Teams — that's fine
      isInTeams = false;
    }
  }

  // --- Auth: Get access token via Teams popup ---
  // mode: 'silent' = no popup, 'post' = Chat.ReadWrite only, 'full' = all scopes including ChatMember.Read
  function getAccessToken(mode) {
    if (!isInTeams) {
      return Promise.reject(new Error('Not running inside Teams.'));
    }

    var authUrl = window.location.origin + '/willbot/auth.html';
    var params = [];
    if (mode === 'silent') params.push('silent=true');
    if (mode === 'full') params.push('full=true');
    if (params.length > 0) authUrl += '?' + params.join('&');

    return microsoftTeams.authentication.authenticate({
      url: authUrl,
      width: 600,
      height: 600
    });
  }

  // Try silent auth on tab load (no popup visible to user)
  function trySilentAuth() {
    if (!isInTeams) return;
    getAccessToken('silent').then(function (token) {
      accessToken = token;
      loadStatus.textContent = 'Ready to post';
      loadStatus.className = 'load-status success';
    }).catch(function () {
      // Silent auth failed — will consent on first spin
    });
  }

  // --- Graph API: Fetch chat members ---
  function fetchChatMembers() {
    loadChatBtn.disabled = true;
    loadStatus.textContent = 'Authenticating...';
    loadStatus.className = 'load-status';

    var chatId = null;
    if (teamsContext && teamsContext.chat && teamsContext.chat.id) {
      chatId = teamsContext.chat.id;
    }

    if (!chatId) {
      loadStatus.textContent = 'No chat context. Open this as a tab in a group chat.';
      loadStatus.className = 'load-status error';
      loadChatBtn.disabled = false;
      return;
    }

    getAccessToken('full').then(function (token) {
      accessToken = token; // store for later use (e.g. posting to chat)
      loadStatus.textContent = 'Loading members...';

      return fetch('https://graph.microsoft.com/v1.0/chats/' + chatId + '/members', {
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json'
        }
      }).then(function (response) {
        if (!response.ok) {
          throw new Error('Graph API error: ' + response.status);
        }
        return response.json();
      }).then(function (data) {
        var members = data.value || [];
        var memberNames = members
          .map(function (m) { return m.displayName || ''; })
          .filter(function (n) { return n.length > 0; });

        if (memberNames.length === 0) {
          loadStatus.textContent = 'No members found.';
          loadStatus.className = 'load-status error';
        } else {
          namesInput.value = memberNames.join('\n');
          updateNames();
          loadStatus.textContent = 'Loaded ' + memberNames.length + ' members!';
          loadStatus.className = 'load-status success';
        }
      });
    }).catch(function (err) {
      console.error('Failed to load chat members:', err);
      loadStatus.textContent = 'Failed: ' + (err.message || err || 'Auth error');
      loadStatus.className = 'load-status error';
    }).finally(function () {
      loadChatBtn.disabled = false;
    });
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

  // --- Post message to chat ---
  function setPostStatus(text, className) {
    if (postStatus) {
      postStatus.textContent = text;
      postStatus.className = 'post-status' + (className ? ' ' + className : '');
    }
  }

  function postMessageToChat(winnerName) {
    setPostStatus('', '');

    if (!postToChatCheckbox || !postToChatCheckbox.checked) {
      return;
    }

    if (!isInTeams) {
      setPostStatus('Not in Teams — message not posted', 'post-error');
      return;
    }

    if (!teamsContext || !teamsContext.chat || !teamsContext.chat.id) {
      setPostStatus('No chat context — open as tab in a group chat', 'post-error');
      return;
    }

    var chatId = teamsContext.chat.id;

    function sendMessage(token) {
      setPostStatus('Posting to chat...', 'posting');

      var messageBody = {
        body: {
          contentType: 'html',
          content: '<b>📖 Next meeting\'s Storyteller: ' + winnerName + '</b><br><i>Get ready to share a fun or interesting story!</i>'
        }
      };

      fetch('https://graph.microsoft.com/v1.0/chats/' + chatId + '/messages', {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(messageBody)
      }).then(function (response) {
        if (!response.ok) {
          return response.json().then(function (err) {
            var msg = (err && err.error && err.error.message) || ('Error ' + response.status);
            setPostStatus('Post failed: ' + msg, 'post-error');
          });
        }
        setPostStatus('Posted to chat!', 'posted');
      }).catch(function (err) {
        setPostStatus('Post failed: ' + (err.message || err), 'post-error');
      });
    }

    if (accessToken) {
      sendMessage(accessToken);
    } else {
      // Try silent first, then consent popup (Chat.ReadWrite only — no admin needed)
      setPostStatus('Connecting...', 'posting');
      getAccessToken('silent').then(function (token) {
        accessToken = token;
        sendMessage(token);
      }).catch(function () {
        // Silent failed — try with consent popup (one-time, no admin needed)
        setPostStatus('Please consent to post...', 'posting');
        getAccessToken('post').then(function (token) {
          accessToken = token;
          sendMessage(token);
        }).catch(function (err) {
          setPostStatus('Auth cancelled — uncheck "Post to chat" or try again', 'post-error');
        });
      });
    }
  }

  function announceWinner() {
    // The pointer is at the top, which is -PI/2 in canvas coordinates.
    // Segment i is drawn from (currentAngle + i * sliceAngle) to (currentAngle + (i+1) * sliceAngle).
    // We need to find which segment contains the angle -PI/2.
    var sliceAngle = (2 * Math.PI) / names.length;
    // Normalize (-PI/2 - currentAngle) into [0, 2PI)
    var pointerAngle = ((-Math.PI / 2 - currentAngle) % (2 * Math.PI) + 2 * Math.PI) % (2 * Math.PI);

    var winnerIndex = Math.floor(pointerAngle / sliceAngle) % names.length;
    var winnerName = names[winnerIndex];

    winnerNameEl.textContent = winnerName;
    winnerOverlay.classList.remove('hidden');
    spawnConfetti();

    // Post result to chat
    postMessageToChat(winnerName);
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
  loadChatBtn.addEventListener('click', fetchChatMembers);
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
