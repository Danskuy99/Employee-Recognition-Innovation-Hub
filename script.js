document.addEventListener('DOMContentLoaded', () => {

  // 1. DUA LIMA AKUN DEMO BAWAAN (Semua Password: 123)
  const defaultAccounts = {
    "william": { username: "william", password: "123", nama: "William", role: "Admin", poin: 999 },
    "dandi": { username: "dandi", password: "123", nama: "Dr. Dandi Siregar, Sp.JP", role: "Dokter", poin: 150 },
    "bryan": { username: "bryan", password: "123", nama: "Dr. Bryan, Sp.JP", role: "Dokter", poin: 120 },
    "wulan": { username: "wulan", password: "123", nama: "Wulan Margi, S.Kep (Nurse)", role: "Perawat", poin: 200 },
    "zhevo": { username: "zhevo", password: "123", nama: "Zhevo Wijaya (Staff)", role: "Staff", poin: 100 }
  };

  const defaultRatings = [
    { penilai: "Dr. Dandi Siregar, Sp.JP", target: "Wulan Margi, S.Kep (Nurse)", score: "5", note: "Sangat cepat dan sigap dalam penanganan darurat jantung." }
  ];

  const defaultKudos = [
    { pengirim: "Dr. Dandi Siregar, Sp.JP", penerima: "Wulan Margi, S.Kep (Nurse)", pesan: "Terima kasih perawat Wulan sudah sangat membantu di ruang tindakan!" }
  ];

  const defaultInovasi = [
    { pengusul: "Zhevo Wijaya (Staff)", judul: "Digitalisasi Form Check-in Pasien", desk: "Mempercepat antrean administrasi poli jantung." }
  ];

  if (!localStorage.getItem('accountsData')) {
    localStorage.setItem('accountsData', JSON.stringify(defaultAccounts));
  }

  let accounts = JSON.parse(localStorage.getItem('accountsData')) || defaultAccounts;
  let ratingsData = JSON.parse(localStorage.getItem('ratingsData')) || defaultRatings;
  let kudosData = JSON.parse(localStorage.getItem('kudosData')) || defaultKudos;
  let inovasiData = JSON.parse(localStorage.getItem('inovasiData')) || defaultInovasi;
  let directChats = JSON.parse(localStorage.getItem('directChatsData')) || [];
  let currentUsername = localStorage.getItem('currentUsername') || null;
  let activeChatTarget = null;
  let activeTabNav = 'navDashboard';

  // DOM ELEMENTS
  const loginOverlay = document.getElementById('loginOverlay');
  const appContainer = document.getElementById('appContainer');
  const adminNavTab = document.getElementById('adminNavTab');
  const userNamaDisplay = document.getElementById('userNamaDisplay');
  const userRoleBadge = document.getElementById('userRoleBadge');
  const userPoinDisplay = document.getElementById('userPoinDisplay');
  const totalUsersCount = document.getElementById('totalUsersCount');
  
  const kudosTargetSelect = document.getElementById('kudosTarget');
  const rateTargetSelect = document.getElementById('rateTargetSelect');

  const chatUsersList = document.getElementById('chatUsersList');
  const activeChatName = document.getElementById('activeChatName');
  const activeChatRole = document.getElementById('activeChatRole');
  const directChatContainer = document.getElementById('directChatContainer');
  const directChatForm = document.getElementById('directChatForm');
  const directChatMessageInput = document.getElementById('directChatMessage');

  const settingsOverlay = document.getElementById('settingsOverlay');
  const openSettingsBtn = document.getElementById('openSettingsBtn');
  const closeSettingsModal = document.getElementById('closeSettingsModal');

  // EXCEL IMPORT & TEMPLATE
  const downloadTemplateBtn = document.getElementById('downloadTemplateBtn');
  if (downloadTemplateBtn) {
    downloadTemplateBtn.addEventListener('click', () => {
      const templateData = [
        { Username: "siti_n", Password: "123", Nama: "Ns. Siti Nurhaliza", Role: "Perawat", Poin: 50 },
        { Username: "budi_s", Password: "123", Nama: "Budi Santoso", Role: "Staff", Poin: 20 }
      ];
      const worksheet = XLSX.utils.json_to_sheet(templateData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Template Pegawai");
      XLSX.writeFile(workbook, "Template_Import_Pegawai_HeartCare.xlsx");
    });
  }

  const excelFileInput = document.getElementById('excelFileInput');
  if (excelFileInput) {
    excelFileInput.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const worksheet = workbook.Sheets[workbook.SheetNames[0]];
          const importedUsers = XLSX.utils.sheet_to_json(worksheet);

          let addedCount = 0;
          importedUsers.forEach(row => {
            const uname = row.Username || row.username;
            const pass = row.Password || row.password || "123";
            const nama = row.Nama || row.nama;
            const role = row.Role || row.role || "Staff";
            const poin = row.Poin || row.poin || 0;

            if (uname && nama) {
              accounts[uname] = {
                username: String(uname).trim(),
                password: String(pass).trim(),
                nama: String(nama).trim(),
                role: String(role).trim(),
                poin: Number(poin) || 0
              };
              addedCount++;
            }
          });

          localStorage.setItem('accountsData', JSON.stringify(accounts));
          Swal.fire({ icon: 'success', title: 'Import Excel Berhasil!', text: `Berhasil menambahkan ${addedCount} pegawai baru.` });
          excelFileInput.value = '';
          updateState();
        } catch (error) {
          Swal.fire({ icon: 'error', title: 'Gagal Membaca File', text: 'Format Excel tidak sesuai template.' });
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  // TAB NAVIGATION
  document.querySelectorAll('.nav-tab-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      switchTab(e.currentTarget.dataset.nav);
    });
  });

  function resetToDefaultTab() {
    switchTab('navDashboard');
  }

  function switchTab(targetNavId) {
    activeTabNav = targetNavId;
    document.querySelectorAll('.nav-tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.nav-section').forEach(s => s.style.display = 'none');

    const activeBtn = document.querySelector(`[data-nav="${targetNavId}"]`);
    const targetElement = document.getElementById(targetNavId);

    if (activeBtn) activeBtn.classList.add('active');
    if (targetElement) {
      targetElement.style.display = targetNavId === 'navChat' ? 'block' : 'grid';
    }

    // Jika masuk ke tab chat dan sedang pilih kontak, tandai pesan dari kontak itu sudah dibaca
    if (targetNavId === 'navChat' && activeChatTarget) {
      markChatAsRead(activeChatTarget);
    }
    checkUnreadMessages();
  }

  function updateState() {
    if (currentUsername && accounts[currentUsername]) {
      loginOverlay.style.display = 'none';
      appContainer.style.display = 'block';

      const user = accounts[currentUsername];
      userNamaDisplay.textContent = user.nama;
      userRoleBadge.textContent = user.role;
      userPoinDisplay.innerHTML = `<i class="fa-solid fa-heart heartbeat-icon"></i> ${user.poin} Poin`;

      if (adminNavTab) adminNavTab.style.display = user.role === 'Admin' ? 'inline-flex' : 'none';
      if (totalUsersCount) totalUsersCount.textContent = Object.keys(accounts).length;
      
      populateDropdowns();
      renderAdminUserList();
      renderChatSidebar();
      renderDirectMessages();
      renderFeed();
      renderLeaderboards();
      checkUnreadMessages();
    } else {
      loginOverlay.style.display = 'flex';
      appContainer.style.display = 'none';
    }
  }

  // SYSTEM NOTIFIKASI CHAT (REALTIME POLLING & BADGE)
  function markChatAsRead(fromUsername) {
    let updated = false;
    directChats.forEach(chat => {
      if (chat.from === fromUsername && chat.to === currentUsername && !chat.read) {
        chat.read = true;
        updated = true;
      }
    });

    if (updated) {
      localStorage.setItem('directChatsData', JSON.stringify(directChats));
      checkUnreadMessages();
      renderChatSidebar();
    }
  }

  function checkUnreadMessages() {
    if (!currentUsername) return;

    // Hitung berapa total pesan belum dibaca untuk user yang sedang login
    const unreadCount = directChats.filter(chat => chat.to === currentUsername && !chat.read).length;
    const chatTabBtn = document.querySelector('[data-nav="navChat"]');

    if (chatTabBtn) {
      let badge = chatTabBtn.querySelector('.chat-unread-badge');
      if (unreadCount > 0) {
        if (!badge) {
          badge = document.createElement('span');
          badge.className = 'chat-unread-badge';
          badge.style.cssText = 'background:#ef4444; color:#fff; border-radius:10px; padding:2px 7px; font-size:0.75rem; margin-left:6px; font-weight:bold; animation: pulse 1.5s infinite;';
          chatTabBtn.appendChild(badge);
        }
        badge.textContent = unreadCount;
      } else if (badge) {
        badge.remove();
      }
    }
  }

  // DETEKSI PESAN MASUK REAL-TIME
  let lastCheckedChatLength = directChats.length;
  setInterval(() => {
    const latestChats = JSON.parse(localStorage.getItem('directChatsData')) || [];
    
    if (latestChats.length > lastCheckedChatLength) {
      const newChat = latestChats[latestChats.length - 1];
      directChats = latestChats;
      lastCheckedChatLength = latestChats.length;

      // Jika pesan baru ditujukan untuk user ini
      if (newChat && newChat.to === currentUsername) {
        const senderName = accounts[newChat.from]?.nama || newChat.from;

        // Tampilkan Notifikasi Toast Pop-up
        Swal.fire({
          toast: true,
          position: 'top-end',
          icon: 'info',
          title: `Pesan baru dari ${senderName}`,
          text: newChat.text.length > 30 ? newChat.text.substring(0, 30) + '...' : newChat.text,
          showConfirmButton: false,
          timer: 3500,
          timerProgressBar: true
        });

        // Tanda pesan terbuka jika berada di room chat ybs
        if (activeTabNav === 'navChat' && activeChatTarget === newChat.from) {
          markChatAsRead(newChat.from);
          renderDirectMessages();
        } else {
          checkUnreadMessages();
          renderChatSidebar();
        }
      }
    }
  }, 1000);

  // FEED RENDERERS
  function renderFeed() {
    const ratingFeedContainer = document.getElementById('ratingFeedList');
    if (ratingFeedContainer) {
      if (ratingsData.length === 0) {
        ratingFeedContainer.innerHTML = `<p class="card-sub" style="text-align:center;">Belum ada penilaian resmi.</p>`;
      } else {
        ratingFeedContainer.innerHTML = ratingsData.map(r => `
          <div class="rating-item" style="background:#fff; padding:12px 16px; border-radius:8px; margin-bottom:10px; border-left:4px solid #e11d48; box-shadow:0 1px 3px rgba(0,0,0,0.05);">
            <div class="rating-item-header" style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
              <span style="font-size:0.9rem;"><strong>${r.penilai}</strong> <i class="fa-solid fa-arrow-right" style="font-size:0.75rem; color:#94a3b8; margin:0 4px;"></i> <strong>${r.target}</strong></span>
              <span class="rating-stars" style="color:#f59e0b; font-size:0.85rem;">${'⭐'.repeat(Number(r.score))}</span>
            </div>
            <p class="card-sub" style="margin:0; font-style:italic; color:#475569; font-size:0.85rem;">"${r.note}"</p>
          </div>
        `).join('');
      }
    }

    const kudosContainer = document.getElementById('kudosFeed');
    const currentUserObj = accounts[currentUsername];
    if (kudosContainer) {
      kudosContainer.innerHTML = kudosData.map(k => {
        const isMine = currentUserObj && k.pengirim === currentUserObj.nama;
        return `
          <div class="chat-bubble-wrapper ${isMine ? 'mine' : 'others'}">
            <div class="chat-meta">${k.pengirim} ➔ ${k.penerima}</div>
            <div class="chat-bubble">${k.pesan}</div>
          </div>
        `;
      }).join('');
    }

    const inovasiContainer = document.getElementById('inovasiFeed');
    if (inovasiContainer) {
      inovasiContainer.innerHTML = inovasiData.map(i => `
        <div class="inovasi-item">
          <h4>${i.judul}</h4>
          <p>${i.desk} — <strong>(${i.pengusul})</strong></p>
        </div>
      `).join('');
    }
  }

  // SEARCH & ADMIN MANAGEMENTS
  const adminSearchUserInput = document.getElementById('adminSearchUserInput');
  if (adminSearchUserInput) {
    adminSearchUserInput.addEventListener('input', () => renderAdminUserList());
  }

  function renderAdminUserList() {
    const adminUserListContainer = document.getElementById('adminUserList');
    if (!adminUserListContainer) return;

    const filterKeyword = adminSearchUserInput ? adminSearchUserInput.value.toLowerCase() : '';
    const filteredUsers = Object.keys(accounts).filter(uname => {
      const u = accounts[uname];
      return u.nama.toLowerCase().includes(filterKeyword) || u.username.toLowerCase().includes(filterKeyword);
    });

    if (filteredUsers.length === 0) {
      adminUserListContainer.innerHTML = `<div style="padding:15px; text-align:center; color:gray;">Tidak ada pegawai ditemukan.</div>`;
      return;
    }

    adminUserListContainer.innerHTML = filteredUsers.map(uname => {
      const u = accounts[uname];
      const isSelf = uname === currentUsername;
      return `
        <div class="admin-user-item">
          <div>
            <strong>${u.nama}</strong> <small>(@${u.username})</small>
            <br/><span class="user-role-badge" style="background:#e2e8f0; color:#334155;">${u.role}</span> • ${u.poin} Poin
          </div>
          ${!isSelf ? `
            <button class="btn-delete-user" data-username="${u.username}">
              <i class="fa-solid fa-trash"></i> Hapus
            </button>
          ` : '<small style="color:gray;">(Akun Anda)</small>'}
        </div>
      `;
    }).join('');

    document.querySelectorAll('.btn-delete-user').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const targetUser = e.currentTarget.dataset.username;
        const targetNama = accounts[targetUser]?.nama || targetUser;

        Swal.fire({
          title: `Hapus Akun ${targetNama}?`,
          text: "Akun ini akan dihapus permanen!",
          icon: 'warning',
          showCancelButton: true,
          confirmButtonColor: '#ef4444',
          cancelButtonColor: '#6b7280',
          confirmButtonText: 'Ya, Hapus!'
        }).then((result) => {
          if (result.isConfirmed) {
            delete accounts[targetUser];
            localStorage.setItem('accountsData', JSON.stringify(accounts));
            Swal.fire('Terhapus!', `Akun ${targetNama} berhasil dihapus.`, 'success');
            updateState();
          }
        });
      });
    });
  }

  function populateDropdowns() {
    if (!kudosTargetSelect || !rateTargetSelect) return;

    kudosTargetSelect.innerHTML = '<option value="">-- Pilih Penerima --</option>';
    rateTargetSelect.innerHTML = '<option value="">-- Pilih Rekan Pegawai --</option>';

    Object.keys(accounts).forEach(uname => {
      if (uname !== currentUsername) {
        const acc = accounts[uname];
        
        const opt1 = document.createElement('option');
        opt1.value = acc.nama; opt1.textContent = `${acc.nama} (${acc.role})`;
        kudosTargetSelect.appendChild(opt1);

        const opt2 = document.createElement('option');
        opt2.value = acc.nama; opt2.textContent = `${acc.nama} (${acc.role})`;
        rateTargetSelect.appendChild(opt2);
      }
    });

    setupSearchableSelect('kudosSearchInput', 'kudosTarget');
    setupSearchableSelect('rateSearchInput', 'rateTargetSelect');
  }

  function setupSearchableSelect(inputId, selectId) {
    const searchInput = document.getElementById(inputId);
    const selectElem = document.getElementById(selectId);
    if (!searchInput || !selectElem) return;

    searchInput.addEventListener('input', (e) => {
      const keyword = e.target.value.toLowerCase();
      let matchFound = false;

      Array.from(selectElem.options).forEach((opt, index) => {
        if (index === 0) return;
        const text = opt.textContent.toLowerCase();
        if (text.includes(keyword)) {
          opt.style.display = '';
          if (!matchFound) {
            selectElem.selectedIndex = index;
            matchFound = true;
          }
        } else {
          opt.style.display = 'none';
        }
      });
    });
  }

  function renderLeaderboards() {
    const allUsers = Object.values(accounts);

    const topStaff = allUsers
      .filter(u => u.role === 'Perawat' || u.role === 'Staff')
      .sort((a, b) => b.poin - a.poin)
      .slice(0, 5);

    const topDoctors = allUsers
      .filter(u => u.role === 'Dokter')
      .sort((a, b) => b.poin - a.poin)
      .slice(0, 5);

    const topStaffContainer = document.getElementById('topStaffList');
    if (topStaffContainer) {
      topStaffContainer.innerHTML = topStaff.length ? topStaff.map((u, i) => `
        <div class="rank-item rank-${i+1}">
          <span class="rank-badge">${i+1}</span>
          <div class="rank-user-info">
            <span>${u.nama}</span> <small>(${u.role})</small>
          </div>
          <span class="rank-points">${u.poin} Poin</span>
        </div>
      `).join('') : '<p class="card-sub">Belum ada data</p>';
    }

    const topDoctorContainer = document.getElementById('topDoctorList');
    if (topDoctorContainer) {
      topDoctorContainer.innerHTML = topDoctors.length ? topDoctors.map((u, i) => `
        <div class="rank-item rank-${i+1}">
          <span class="rank-badge">${i+1}</span>
          <div class="rank-user-info">
            <span>${u.nama}</span>
          </div>
          <span class="rank-points">${u.poin} Poin</span>
        </div>
      `).join('') : '<p class="card-sub">Belum ada data</p>';
    }
  }

  // CHAT CONTACTS & MESSAGES
  const chatSearchContact = document.getElementById('chatSearchContact');
  if (chatSearchContact) {
    chatSearchContact.addEventListener('input', () => renderChatSidebar());
  }

  function renderChatSidebar() {
    if (!chatUsersList) return;
    chatUsersList.innerHTML = '';

    const filterKey = chatSearchContact ? chatSearchContact.value.toLowerCase() : '';

    Object.keys(accounts).forEach(uname => {
      if (uname !== currentUsername) {
        const acc = accounts[uname];
        if (acc.nama.toLowerCase().includes(filterKey) || acc.role.toLowerCase().includes(filterKey)) {
          const isActive = activeChatTarget === uname ? 'active' : '';

          // Cek apakah ada pesan belum dibaca dari pengguna ini
          const hasUnread = directChats.some(chat => chat.from === uname && chat.to === currentUsername && !chat.read);

          const userItem = document.createElement('div');
          userItem.className = `chat-user-item ${isActive}`;
          userItem.style.position = 'relative';
          userItem.innerHTML = `
            <i class="fa-solid fa-circle-user chat-user-avatar"></i>
            <div class="chat-user-details" style="flex:1;">
              <h4 style="display:flex; justify-content:space-between; align-items:center;">
                ${acc.nama}
                ${hasUnread ? '<span style="width:8px; height:8px; background:#ef4444; border-radius:50%; display:inline-block;"></span>' : ''}
              </h4>
              <p>@${acc.username} • ${acc.role}</p>
            </div>
          `;

          userItem.addEventListener('click', () => {
            activeChatTarget = uname;
            markChatAsRead(uname);
            renderChatSidebar();
            renderDirectMessages();
          });

          chatUsersList.appendChild(userItem);
        }
      }
    });
  }

  function renderDirectMessages() {
    if (!activeChatTarget || !accounts[activeChatTarget]) {
      activeChatName.textContent = "Pilih Kontak Chat";
      activeChatRole.textContent = "Internal Chat System";
      directChatContainer.innerHTML = `
        <div class="chat-placeholder">
          <i class="fa-solid fa-comments"></i>
          <p>Pilih salah satu teman kerja untuk mulai berdiskusi!</p>
        </div>
      `;
      directChatForm.style.display = 'none';
      return;
    }

    const targetUser = accounts[activeChatTarget];
    activeChatName.textContent = targetUser.nama;
    activeChatRole.textContent = `@${targetUser.username} • ${targetUser.role}`;
    directChatForm.style.display = 'flex';

    const filteredMessages = directChats.filter(chat => 
      (chat.from === currentUsername && chat.to === activeChatTarget) ||
      (chat.from === activeChatTarget && chat.to === currentUsername)
    );

    if (filteredMessages.length === 0) {
      directChatContainer.innerHTML = `
        <div class="chat-placeholder">
          <i class="fa-solid fa-paper-plane"></i>
          <p>Belum ada obrolan dengan <strong>${targetUser.nama}</strong>.</p>
        </div>
      `;
    } else {
      directChatContainer.innerHTML = filteredMessages.map(msg => {
        const isSentByMe = msg.from === currentUsername;
        return `
          <div class="dm-bubble ${isSentByMe ? 'sent' : 'received'}">
            <div>${msg.text}</div>
            <div class="dm-time">${msg.time}</div>
          </div>
        `;
      }).join('');
      directChatContainer.scrollTop = directChatContainer.scrollHeight;
    }
  }

  if (directChatForm) {
    directChatForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const text = directChatMessageInput.value.trim();
      if (!activeChatTarget || !text) return;

      const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      
      // Kirim pesan dengan status awal 'read: false' untuk penerima
      directChats.push({ 
        from: currentUsername, 
        to: activeChatTarget, 
        text, 
        time,
        read: false 
      });

      localStorage.setItem('directChatsData', JSON.stringify(directChats));
      lastCheckedChatLength = directChats.length;
      directChatMessageInput.value = '';
      renderDirectMessages();
    });
  }

  // AUTH LOGIC
  document.getElementById('loginForm').addEventListener('submit', (e) => {
    e.preventDefault();
    const uname = document.getElementById('loginUser').value.trim();
    const pass = document.getElementById('loginPass').value.trim();

    accounts = JSON.parse(localStorage.getItem('accountsData')) || defaultAccounts;

    if (accounts[uname] && String(accounts[uname].password) === pass) {
      currentUsername = uname;
      localStorage.setItem('currentUsername', currentUsername);
      resetToDefaultTab();

      Swal.fire({
        icon: 'success',
        title: 'Login Berhasil! 🎉',
        text: `Selamat datang, ${accounts[uname].nama}`,
        timer: 1500,
        showConfirmButton: false
      }).then(() => updateState());
    } else {
      Swal.fire({
        icon: 'error',
        title: 'Akses Ditolak',
        text: 'Username atau Password salah!',
        confirmButtonColor: '#e11d48'
      });
    }
  });

  document.getElementById('resetDataBtn').addEventListener('click', () => {
    localStorage.clear();
    accounts = defaultAccounts;
    localStorage.setItem('accountsData', JSON.stringify(defaultAccounts));
    Swal.fire({ icon: 'info', title: 'Data Di-reset!', text: '5 Akun Demo dikembalikan (Password: 123)', timer: 1500, showConfirmButton: false });
    resetToDefaultTab();
    updateState();
  });

  document.getElementById('logoutBtn').addEventListener('click', () => {
    Swal.fire({
      title: 'Keluar Portal?', icon: 'warning', showCancelButton: true,
      confirmButtonColor: '#e11d48', confirmButtonText: 'Ya, Logout'
    }).then((result) => {
      if (result.isConfirmed) {
        localStorage.removeItem('currentUsername');
        currentUsername = null;
        activeChatTarget = null;
        resetToDefaultTab();
        updateState();
      }
    });
  });

  // SUBMIT FORM EVALUASI, KUDOS, INOVASI
  const generalForm = document.getElementById('generalRatingForm');
  if (generalForm) {
    generalForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const targetNama = rateTargetSelect.value;
      const score = document.getElementById('rateScoreSelect').value;
      const note = document.getElementById('rateNoteText').value;

      ratingsData.unshift({ penilai: accounts[currentUsername].nama, target: targetNama, score, note });
      localStorage.setItem('ratingsData', JSON.stringify(ratingsData));

      const targetUname = Object.keys(accounts).find(k => accounts[k].nama === targetNama);
      if (targetUname) accounts[targetUname].poin += (Number(score) * 5);
      localStorage.setItem('accountsData', JSON.stringify(accounts));

      Swal.fire({ icon: 'success', title: 'Evaluasi Terkirim!', text: `Terima kasih atas masukan untuk ${targetNama}!`, confirmButtonColor: '#e11d48' });
      generalForm.reset();
      updateState();
    });
  }

  document.getElementById('kudosForm').addEventListener('submit', (e) => {
    e.preventDefault();
    const penerima = kudosTargetSelect.value;
    const pesan = document.getElementById('kudosPesan').value;

    const sender = accounts[currentUsername];
    kudosData.unshift({ pengirim: sender.nama, penerima, pesan });
    sender.poin += 10;

    localStorage.setItem('kudosData', JSON.stringify(kudosData));
    localStorage.setItem('accountsData', JSON.stringify(accounts));

    Swal.fire({ icon: 'success', title: 'Apresiasi Terkirim!', text: '+10 Poin Apresiasi dikreditkan!', timer: 1800, showConfirmButton: false });
    document.getElementById('kudosForm').reset();
    updateState();
  });

  document.getElementById('inovasiForm').addEventListener('submit', (e) => {
    e.preventDefault();
    const judul = document.getElementById('inovasiJudul').value;
    const desk = document.getElementById('inovasiDesk').value;

    inovasiData.unshift({ pengusul: accounts[currentUsername].nama, judul, desk });
    localStorage.setItem('inovasiData', JSON.stringify(inovasiData));

    Swal.fire({ icon: 'success', title: 'Ide Terkirim!', text: 'Terima kasih atas ide inovasimu!', timer: 1800, showConfirmButton: false });
    document.getElementById('inovasiForm').reset();
    updateState();
  });

  // SETTINGS MODAL
  if (openSettingsBtn) {
    openSettingsBtn.addEventListener('click', () => {
      const user = accounts[currentUsername];
      document.getElementById('profileNameInput').value = user.nama;
      document.getElementById('profileRoleInput').value = user.role;
      
      document.querySelectorAll('.settings-tabs .tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.style.display = 'none');
      document.querySelector('[data-tab="tabProfile"]').classList.add('active');
      document.getElementById('tabProfile').style.display = 'block';

      settingsOverlay.style.display = 'flex';
    });
  }
  if (closeSettingsModal) closeSettingsModal.addEventListener('click', () => settingsOverlay.style.display = 'none');

  resetToDefaultTab();
  updateState();
});