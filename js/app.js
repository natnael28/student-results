// ===== Storage =====
const Storage = {
  _get(key) {
    try { return JSON.parse(localStorage.getItem(key)) || []; }
    catch { return []; }
  },
  _set(key, data) {
    localStorage.setItem(key, JSON.stringify(data));
  },
  getUsers() { return this._get('sr_users'); },
  setUsers(u) { this._set('sr_users', u); },
  getResults() { return this._get('sr_results'); },
  setResults(r) { this._set('sr_results', r); },
  getCurrentUser() {
    try { return JSON.parse(sessionStorage.getItem('sr_current')); }
    catch { return null; }
  },
  setCurrentUser(u) {
    if (u) sessionStorage.setItem('sr_current', JSON.stringify(u));
    else sessionStorage.removeItem('sr_current');
  },
  getGradeScale() {
    const def = [
      { min: 90, max: 100, letter: 'A+', gpa: 4.0 },
      { min: 85, max: 89, letter: 'A', gpa: 4.0 },
      { min: 80, max: 84, letter: 'A-', gpa: 3.75 },
      { min: 75, max: 79, letter: 'B+', gpa: 3.5 },
      { min: 70, max: 74, letter: 'B', gpa: 3.0 },
      { min: 65, max: 69, letter: 'B-', gpa: 2.75 },
      { min: 60, max: 64, letter: 'C+', gpa: 2.5 },
      { min: 50, max: 59, letter: 'C', gpa: 2.0 },
      { min: 45, max: 49, letter: 'D', gpa: 1.0 },
      { min: 0, max: 44, letter: 'F', gpa: 0 }
    ];
    try {
      const saved = JSON.parse(localStorage.getItem('sr_grade_scale'));
      if (saved && saved.length) return saved;
    } catch {}
    return def;
  },
  setGradeScale(s) { localStorage.setItem('sr_grade_scale', JSON.stringify(s)); },
  getSubjects() { return this._get('sr_subjects'); },
  setSubjects(s) { this._set('sr_subjects', s); },
  seedInstructor() {
    const users = this.getUsers();
    if (!users.find(u => u.email === 'admin@srms.com')) {
      users.push({
        id: Date.now().toString(36),
        name: 'Instructor',
        email: 'admin@srms.com',
        password: 'admin123',
        role: 'instructor',
        profile: { phone: '', address: '' }
      });
      this.setUsers(users);
    }
    if (this.getSubjects().length === 0) {
      this.setSubjects(['Mathematics', 'English', 'Science']);
    }
  }
};

// ===== Auth =====
const Auth = {
  currentUser: null,

  init() {
    Storage.seedInstructor();
    this.currentUser = Storage.getCurrentUser();
  },

  login(identifier, password) {
    const users = Storage.getUsers();
    const isEmail = identifier.includes('@');
    const user = isEmail
      ? users.find(u => u.email === identifier.toLowerCase() && u.password === password)
      : users.find(u => u.studentId === identifier && u.password === password);
    if (!user) return { ok: false, msg: 'Invalid credentials.' };
    this.currentUser = user;
    Storage.setCurrentUser(user);
    return { ok: true };
  },

  register(firstName, fatherName, studentId, password) {
    const users = Storage.getUsers();
    if (users.find(u => u.studentId === studentId))
      return { ok: false, msg: 'Student ID already registered.' };
    const user = {
      id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      firstName, fatherName, studentId, password, role: 'student',
      profile: { phone: '', address: '' }
    };
    users.push(user);
    Storage.setUsers(users);
    return { ok: true, msg: 'Registration successful! Please login.' };
  },

  logout() {
    this.currentUser = null;
    Storage.setCurrentUser(null);
  },

  updateProfile(data) {
    if (!this.currentUser) return;
    const users = Storage.getUsers();
    const idx = users.findIndex(u => u.id === this.currentUser.id);
    if (idx === -1) return;
    const keys = Object.keys(data);
    keys.forEach(k => { users[idx][k] = data[k]; });
    if (data.profile) {
      users[idx].profile = { ...users[idx].profile, ...data.profile };
    }
    Storage.setUsers(users);
    this.currentUser = users[idx];
    Storage.setCurrentUser(this.currentUser);
  },

  resetPassword(studentId) {
    const users = Storage.getUsers();
    const user = users.find(u => u.studentId === studentId);
    if (!user) return { ok: false, msg: 'No account found with that ID.' };
    const newPwd = 'reset123';
    user.password = newPwd;
    Storage.setUsers(users);
    return { ok: true, msg: `Password reset to "${newPwd}". Please login.` };
  },

  changePassword(oldPwd, newPwd) {
    if (!this.currentUser) return { ok: false, msg: 'Not logged in.' };
    if (this.currentUser.password !== oldPwd)
      return { ok: false, msg: 'Current password is incorrect.' };
    const users = Storage.getUsers();
    const idx = users.findIndex(u => u.id === this.currentUser.id);
    users[idx].password = newPwd;
    this.currentUser.password = newPwd;
    Storage.setUsers(users);
    Storage.setCurrentUser(this.currentUser);
    return { ok: true, msg: 'Password changed successfully.' };
  }
};

// ===== Column name normalization =====
const COL_MAP = {
  'student id': 'sid', 'studentid': 'sid', 'id': 'sid', 'student_id': 'sid',
  'mid': 'mid', 'mid exam': 'mid', 'midterm': 'mid', 'mid_exam': 'mid',
  'quiz': 'quiz', 'quize': 'quiz', 'quizzes': 'quiz',
  'project': 'project', 'projects': 'project',
  'assignment': 'assignment', 'assignments': 'assignment', 'assigment': 'assignment', 'homework': 'assignment', 'hw': 'assignment',
  'final': 'final', 'final exam': 'final', 'finalterm': 'final', 'final_exam': 'final',
  'total': 'total', 'totals': 'total', 'overall': 'total', 'grade': 'total', 'score': 'total'
};

function normalizeCol(name) {
  return COL_MAP[name.toString().trim().toLowerCase().replace(/\s+/g, ' ')] || null;
}

// ===== Results =====
const Results = {
  upload(subject, file, callback) {
    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet);
        if (!json || json.length === 0) {
          callback({ ok: false, msg: 'Excel file is empty.' });
          return;
        }
        const headers = Object.keys(json[0]);
        const colMap = {};
        headers.forEach(h => {
          const key = normalizeCol(h);
          if (key) colMap[key] = h;
        });
        if (!colMap.sid) {
          callback({ ok: false, msg: 'Missing "Student ID" column.' });
          return;
        }
        const scoreKeys = ['mid', 'quiz', 'project', 'assignment', 'final', 'total'];
        const hasScore = scoreKeys.some(k => colMap[k]);
        if (!hasScore) {
          callback({ ok: false, msg: 'No score columns found. Add any of: Mid Exam, Quiz, Project, Assignment, Final Exam, Total.' });
          return;
        }
        const results = Storage.getResults();
        let count = 0;
        json.forEach(row => {
          const sid = row[colMap.sid].toString().trim();
          if (!sid) return;
          let existing = results.findIndex(r => r.studentId === sid && r.subject === subject);
          if (existing === -1) {
            results.push({ studentId: sid, subject: subject, mid: null, quiz: null, project: null, assignment: null, final: null, total: null });
            existing = results.length - 1;
          }
          scoreKeys.forEach(key => {
            if (colMap[key]) {
              const val = parseFloat(row[colMap[key]]);
              if (!isNaN(val)) results[existing][key] = val;
            }
          });
          count++;
        });
        Storage.setResults(results);
        callback({ ok: true, msg: `${count} student results uploaded/updated for ${subject}.` });
      } catch (err) {
        callback({ ok: false, msg: 'Failed to parse Excel file. Check column names.' });
      }
    };
    reader.readAsArrayBuffer(file);
  },

  getByStudent(sid) {
    const all = Storage.getResults();
    return all.filter(r => r.studentId === sid);
  },

  getAll() {
    return Storage.getResults();
  },

  getStudentSummary(sid) {
    const rows = this.getByStudent(sid);
    return rows.map(r => ({
      subject: r.subject,
      mid: r.mid ?? '-',
      quiz: r.quiz ?? '-',
      project: r.project ?? '-',
      assignment: r.assignment ?? '-',
      final: r.final ?? '-',
      total: r.total ?? ((r.mid || 0) + (r.quiz || 0) + (r.project || 0) + (r.assignment || 0) + (r.final || 0))
    }));
  }
};

// ===== Grading =====
function calculateGrade(total) {
  const scale = Storage.getGradeScale();
  for (let i = 0; i < scale.length; i++) {
    if (total >= scale[i].min && total <= scale[i].max) {
      return { letter: scale[i].letter, gpa: scale[i].gpa };
    }
  }
  return { letter: 'F', gpa: 0 };
}

// ===== UI =====
document.addEventListener('DOMContentLoaded', function () {
  Storage.seedInstructor();
  Auth.init();
  render();
});

function render() {
  const user = Auth.currentUser;
  if (user) {
    document.getElementById('authSection').classList.add('hidden');
    document.getElementById('appSection').classList.remove('hidden');
    document.getElementById('navUserName').textContent = user.role === 'instructor' ? user.name : (user.firstName + ' ' + user.fatherName);
    document.getElementById('navUserRole').textContent = user.role === 'instructor' ? 'Instructor' : (user.studentId || '');

    if (user.role === 'instructor') {
      document.getElementById('studentDashboard').classList.add('hidden');
      document.getElementById('instructorDashboard').classList.remove('hidden');
      renderInstructorDashboard();
    } else {
      document.getElementById('instructorDashboard').classList.add('hidden');
      document.getElementById('studentDashboard').classList.remove('hidden');
      renderStudentDashboard();
    }
  } else {
    document.getElementById('authSection').classList.remove('hidden');
    document.getElementById('appSection').classList.add('hidden');
  }
}

// ===== Auth UI =====
function showLogin() {
  document.getElementById('loginForm').classList.remove('hidden');
  document.getElementById('registerForm').classList.add('hidden');
  document.getElementById('resetForm').classList.add('hidden');
  document.getElementById('tabLogin').classList.add('active');
  document.getElementById('tabRegister').classList.remove('active');
  document.getElementById('authMsg').className = 'msg';
  document.getElementById('authMsg').textContent = '';
}

function showRegister() {
  document.getElementById('loginForm').classList.add('hidden');
  document.getElementById('registerForm').classList.remove('hidden');
  document.getElementById('resetForm').classList.add('hidden');
  document.getElementById('tabRegister').classList.add('active');
  document.getElementById('tabLogin').classList.remove('active');
  document.getElementById('authMsg').className = 'msg';
  document.getElementById('authMsg').textContent = '';
}

function showReset() {
  document.getElementById('loginForm').classList.add('hidden');
  document.getElementById('registerForm').classList.add('hidden');
  document.getElementById('resetForm').classList.remove('hidden');
  document.getElementById('authMsg').className = 'msg';
  document.getElementById('authMsg').textContent = '';
}

function handleLogin() {
  const identifier = document.getElementById('loginId').value.trim();
  const password = document.getElementById('loginPassword').value;
  const msgEl = document.getElementById('authMsg');
  if (!identifier || !password) {
    msgEl.className = 'msg error'; msgEl.textContent = 'Please fill all fields.'; return;
  }
  const result = Auth.login(identifier, password);
  if (result.ok) {
    render();
  } else {
    msgEl.className = 'msg error'; msgEl.textContent = result.msg;
  }
}

function handleRegister() {
  const firstName = document.getElementById('regFirstName').value.trim();
  const fatherName = document.getElementById('regFatherName').value.trim();
  const studentId = document.getElementById('regStudentId').value.trim();
  const password = document.getElementById('regPassword').value;
  const confirm = document.getElementById('regConfirm').value;
  const msgEl = document.getElementById('authMsg');
  if (!firstName || !fatherName || !studentId || !password || !confirm) {
    msgEl.className = 'msg error'; msgEl.textContent = 'Please fill all fields.'; return;
  }
  if (password !== confirm) {
    msgEl.className = 'msg error'; msgEl.textContent = 'Passwords do not match.'; return;
  }
  if (password.length < 4) {
    msgEl.className = 'msg error'; msgEl.textContent = 'Password must be at least 4 characters.'; return;
  }
  const result = Auth.register(firstName, fatherName, studentId, password);
  if (result.ok) {
    msgEl.className = 'msg success'; msgEl.textContent = result.msg;
    showLogin();
    document.getElementById('loginId').value = studentId;
  } else {
    msgEl.className = 'msg error'; msgEl.textContent = result.msg;
  }
}

function handleReset() {
  const studentId = document.getElementById('resetId').value.trim();
  const msgEl = document.getElementById('authMsg');
  if (!studentId) {
    msgEl.className = 'msg error'; msgEl.textContent = 'Please enter your Student ID.'; return;
  }
  const result = Auth.resetPassword(studentId);
  msgEl.className = result.ok ? 'msg success' : 'msg error';
  msgEl.textContent = result.msg;
  if (result.ok) {
    setTimeout(() => showLogin(), 2000);
  }
}

function handleLogout() {
  Auth.logout();
  render();
}

// ===== Student Dashboard =====
function renderStudentDashboard() {
  const user = Auth.currentUser;
  if (!user) return;
  document.getElementById('studentName').textContent = user.firstName + ' ' + user.fatherName;
  document.getElementById('studentIdDisplay').textContent = user.studentId;
  document.getElementById('studentPhone').textContent = user.profile.phone || 'Not set';
  document.getElementById('studentAddress').textContent = user.profile.address || 'Not set';

  renderStudentResults(user.studentId);
}

function renderStudentResults(sid) {
  const tbody = document.getElementById('resultsBody');
  const empty = document.getElementById('resultsEmpty');
  const summary = Results.getStudentSummary(sid);
  tbody.innerHTML = '';
  if (summary.length === 0) {
    empty.classList.remove('hidden');
    return;
  }
  empty.classList.add('hidden');
  summary.forEach(r => {
    const grade = r.total !== '-' ? calculateGrade(r.total) : { letter: '-' };
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${r.subject}</td><td>${r.mid}</td><td>${r.quiz}</td><td>${r.project}</td><td>${r.assignment}</td><td>${r.final}</td><td><strong>${r.total}</strong></td><td><strong>${grade.letter}</strong></td>`;
    tbody.appendChild(tr);
  });
}

// Student Profile Modal
function openProfileModal() {
  const user = Auth.currentUser;
  if (!user) return;
  document.getElementById('editFirstName').value = user.firstName || user.name || '';
  document.getElementById('editFatherName').value = user.fatherName || '';
  document.getElementById('editPhone').value = user.profile.phone || '';
  document.getElementById('editAddress').value = user.profile.address || '';
  document.getElementById('profileModal').classList.remove('hidden');
}

function closeProfileModal() {
  document.getElementById('profileModal').classList.add('hidden');
}

function saveProfile() {
  const user = Auth.currentUser;
  if (!user) return;
  const firstName = document.getElementById('editFirstName').value.trim();
  const fatherName = document.getElementById('editFatherName').value.trim();
  const phone = document.getElementById('editPhone').value.trim();
  const address = document.getElementById('editAddress').value.trim();
  if (!firstName) {
    alert('First name is required.');
    return;
  }
  if (user.role === 'instructor') {
    Auth.updateProfile({ name: firstName, profile: { phone, address } });
  } else {
    Auth.updateProfile({ firstName, fatherName, profile: { phone, address } });
  }
  closeProfileModal();
  render();
}

// Change Password Modal
function openPwdModal() {
  document.getElementById('pwdModal').classList.remove('hidden');
  document.getElementById('pwdMsg').className = 'msg';
  document.getElementById('pwdMsg').textContent = '';
  document.getElementById('oldPwd').value = '';
  document.getElementById('newPwd').value = '';
  document.getElementById('confirmPwd').value = '';
}

function closePwdModal() {
  document.getElementById('pwdModal').classList.add('hidden');
}

function changePassword() {
  const oldPwd = document.getElementById('oldPwd').value;
  const newPwd = document.getElementById('newPwd').value;
  const confirmPwd = document.getElementById('confirmPwd').value;
  const msgEl = document.getElementById('pwdMsg');
  if (!oldPwd || !newPwd || !confirmPwd) {
    msgEl.className = 'msg error'; msgEl.textContent = 'Please fill all fields.'; return;
  }
  if (newPwd !== confirmPwd) {
    msgEl.className = 'msg error'; msgEl.textContent = 'New passwords do not match.'; return;
  }
  const result = Auth.changePassword(oldPwd, newPwd);
  msgEl.className = result.ok ? 'msg success' : 'msg error';
  msgEl.textContent = result.msg;
  if (result.ok) setTimeout(() => closePwdModal(), 1500);
}

// ===== Instructor: Register Student =====
function handleInstructorRegisterStudent() {
  const firstName = document.getElementById('instRegFirstName').value.trim();
  const fatherName = document.getElementById('instRegFatherName').value.trim();
  const studentId = document.getElementById('instRegStudentId').value.trim();
  const password = document.getElementById('instRegPassword').value;
  const msgEl = document.getElementById('instRegMsg');
  if (!firstName || !fatherName || !studentId || !password) {
    msgEl.className = 'msg error'; msgEl.textContent = 'Please fill all fields.'; return;
  }
  const result = Auth.register(firstName, fatherName, studentId, password);
  msgEl.className = result.ok ? 'msg success' : 'msg error';
  msgEl.textContent = result.msg;
  if (result.ok) {
    document.getElementById('instRegFirstName').value = '';
    document.getElementById('instRegFatherName').value = '';
    document.getElementById('instRegStudentId').value = '';
    document.getElementById('instRegPassword').value = '';
    renderInstructorDashboard();
  }
}

function handleBulkRegister(input) {
  const msgEl = document.getElementById('instRegMsg');
  if (!input.files || !input.files[0]) return;
  const file = input.files[0];
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet);
      if (!json || json.length === 0) {
        msgEl.className = 'msg error'; msgEl.textContent = 'Excel file is empty.'; return;
      }
      let count = 0;
      json.forEach(row => {
        const firstName = (row['First Name'] || row['firstName'] || row['FirstName'] || row['first name'] || '').toString().trim();
        const fatherName = (row['Father Name'] || row['fatherName'] || row['FatherName'] || row['father name'] || '').toString().trim();
        const studentId = (row['Student ID'] || row['studentId'] || row['StudentId'] || row['ID'] || row['id'] || '').toString().trim();
        const password = (row['Password'] || row['password'] || 'student123').toString().trim();
        if (firstName && fatherName && studentId) {
          const r = Auth.register(firstName, fatherName, studentId, password);
          if (r.ok) count++;
        }
      });
      msgEl.className = count > 0 ? 'msg success' : 'msg error';
      msgEl.textContent = count > 0 ? `${count} students registered successfully.` : 'No valid students found. Check columns: First Name, Father Name, Student ID, Password';
      if (count > 0) renderInstructorDashboard();
    } catch (err) {
      msgEl.className = 'msg error'; msgEl.textContent = 'Failed to parse Excel file.';
    }
  };
  reader.readAsArrayBuffer(file);
  input.value = '';
}

// ===== Subjects =====
function addSubject() {
  const input = document.getElementById('newSubjectInput');
  const name = input.value.trim();
  const msgEl = document.getElementById('subjectsList');
  if (!name) return;
  const subjects = Storage.getSubjects();
  if (subjects.includes(name)) {
    input.value = '';
    return;
  }
  subjects.push(name);
  Storage.setSubjects(subjects);
  input.value = '';
  renderSubjectsList();
  populateSubjectDropdowns();
}

function removeSubject(name) {
  let subjects = Storage.getSubjects();
  subjects = subjects.filter(s => s !== name);
  Storage.setSubjects(subjects);
  const results = Storage.getResults().filter(r => r.subject !== name);
  Storage.setResults(results);
  renderSubjectsList();
  populateSubjectDropdowns();
  renderAllResultsTable();
}

function renderSubjectsList() {
  const el = document.getElementById('subjectsList');
  if (!el) return;
  const subjects = Storage.getSubjects();
  if (subjects.length === 0) {
    el.innerHTML = '<span style="color:#888;font-size:0.85rem;">No subjects added yet.</span>';
    return;
  }
  el.innerHTML = subjects.map(s =>
    `<span style="display:inline-flex;align-items:center;gap:0.3rem;padding:0.25rem 0.6rem;background:#f0f5ff;border-radius:4px;font-size:0.85rem;">${s} <button class="rm-subj" data-subject="${s.replace(/"/g,'&quot;')}" style="background:none;border:none;cursor:pointer;color:#999;font-size:1rem;line-height:1;">&times;</button></span>`
  ).join('');
  el.querySelectorAll('.rm-subj').forEach(btn => {
    btn.onclick = function () { removeSubject(this.dataset.subject); };
  });
}

function populateSubjectDropdowns() {
  const subjects = Storage.getSubjects();
  ['manualSubject', 'uploadSubject'].forEach(id => {
    const sel = document.getElementById(id);
    if (!sel) return;
    const current = sel.value;
    sel.innerHTML = '<option value="">-- Select Subject --</option>';
    subjects.forEach(s => {
      const opt = document.createElement('option');
      opt.value = s;
      opt.textContent = s;
      sel.appendChild(opt);
    });
    if (subjects.includes(current)) sel.value = current;
  });
}

// ===== Manual Entry =====
function populateManualStudentDropdown() {
  const sel = document.getElementById('manualStudent');
  if (!sel) return;
  const current = sel.value;
  sel.innerHTML = '<option value="">-- Select Student --</option>';
  Storage.getUsers().filter(u => u.role === 'student').forEach(u => {
    const opt = document.createElement('option');
    opt.value = u.studentId;
    opt.textContent = `${u.firstName} ${u.fatherName} (${u.studentId})`;
    sel.appendChild(opt);
  });
  sel.value = current;
}

function getManualScores() {
  return {
    mid: parseFloat(document.getElementById('manualMid').value) || 0,
    quiz: parseFloat(document.getElementById('manualQuiz').value) || 0,
    project: parseFloat(document.getElementById('manualProject').value) || 0,
    assignment: parseFloat(document.getElementById('manualAssignment').value) || 0,
    final: parseFloat(document.getElementById('manualFinal').value) || 0
  };
}

function calculateManualEntry() {
  const scores = getManualScores();
  const total = scores.mid + scores.quiz + scores.project + scores.assignment + scores.final;
  const grade = calculateGrade(total);
  document.getElementById('manualCalcResult').textContent = `Total: ${total}  |  Grade: ${grade.letter} (GPA: ${grade.gpa})`;
}

function saveManualEntry() {
  const sid = document.getElementById('manualStudent').value;
  const subject = document.getElementById('manualSubject').value.trim();
  const msgEl = document.getElementById('manualMsg');
  if (!sid) { msgEl.className = 'msg error'; msgEl.textContent = 'Please select a student.'; return; }
  if (!subject) { msgEl.className = 'msg error'; msgEl.textContent = 'Please enter a subject.'; return; }
  const scores = getManualScores();
  const total = scores.mid + scores.quiz + scores.project + scores.assignment + scores.final;
  const results = Storage.getResults();
  let existing = results.findIndex(r => r.studentId === sid && r.subject === subject);
  if (existing === -1) {
    results.push({ studentId: sid, subject, mid: null, quiz: null, project: null, assignment: null, final: null, total: null });
    existing = results.length - 1;
  }
  results[existing].mid = scores.mid;
  results[existing].quiz = scores.quiz;
  results[existing].project = scores.project;
  results[existing].assignment = scores.assignment;
  results[existing].final = scores.final;
  results[existing].total = total;
  Storage.setResults(results);
  msgEl.className = 'msg success';
  const grade = calculateGrade(total);
  msgEl.textContent = `Saved. Total: ${total}, Grade: ${grade.letter}`;
  document.getElementById('manualCalcResult').textContent = '';
  renderAllResultsTable();
}

// ===== Grading Scale UI =====
function renderGradeScaleDisplay() {
  const el = document.getElementById('gradeScaleDisplay');
  if (!el) return;
  const scale = Storage.getGradeScale();
  el.innerHTML = scale.map(s => `<span style="display:inline-block;margin:0.2rem 0.4rem 0.2rem 0;padding:0.15rem 0.5rem;background:#f0f5ff;border-radius:4px;font-size:0.8rem;">${s.min}–${s.max === 100 ? '100' : s.max} → <strong>${s.letter}</strong> (${s.gpa})</span>`).join('');
}

function openGradeScaleModal() {
  const scale = Storage.getGradeScale();
  const form = document.getElementById('gradeScaleForm');
  form.innerHTML = scale.map((s, i) => `
    <div style="display:flex;gap:0.5rem;align-items:center;margin-bottom:0.5rem;">
      <input type="number" id="gs_min_${i}" value="${s.min}" style="width:60px;padding:0.4rem;border:1px solid #d0d5dd;border-radius:4px;font-size:0.85rem;">
      <span>–</span>
      <input type="number" id="gs_max_${i}" value="${s.max}" style="width:60px;padding:0.4rem;border:1px solid #d0d5dd;border-radius:4px;font-size:0.85rem;">
      <span>→</span>
      <input type="text" id="gs_letter_${i}" value="${s.letter}" style="width:40px;padding:0.4rem;border:1px solid #d0d5dd;border-radius:4px;font-size:0.85rem;text-align:center;">
      <span style="font-size:0.8rem;color:#888;">GPA</span>
      <input type="number" id="gs_gpa_${i}" value="${s.gpa}" step="0.25" style="width:65px;padding:0.4rem;border:1px solid #d0d5dd;border-radius:4px;font-size:0.85rem;">
    </div>
  `).join('');
  document.getElementById('gradeScaleMsg').className = 'msg';
  document.getElementById('gradeScaleMsg').textContent = '';
  document.getElementById('gradeScaleModal').classList.remove('hidden');
}

function closeGradeScaleModal() {
  document.getElementById('gradeScaleModal').classList.add('hidden');
}

function saveGradeScale() {
  const scale = Storage.getGradeScale();
  const newScale = [];
  for (let i = 0; i < scale.length; i++) {
    const min = parseInt(document.getElementById(`gs_min_${i}`).value);
    const max = parseInt(document.getElementById(`gs_max_${i}`).value);
    const letter = document.getElementById(`gs_letter_${i}`).value.trim();
    const gpa = parseFloat(document.getElementById(`gs_gpa_${i}`).value);
    if (isNaN(min) || isNaN(max) || !letter || isNaN(gpa)) {
      document.getElementById('gradeScaleMsg').className = 'msg error';
      document.getElementById('gradeScaleMsg').textContent = 'Invalid values in row ' + (i + 1);
      return;
    }
    newScale.push({ min, max, letter, gpa });
  }
  Storage.setGradeScale(newScale);
  document.getElementById('gradeScaleMsg').className = 'msg success';
  document.getElementById('gradeScaleMsg').textContent = 'Grading scale saved.';
  renderGradeScaleDisplay();
  renderAllResultsTable();
  if (Auth.currentUser && Auth.currentUser.role === 'student') {
    renderStudentResults(Auth.currentUser.studentId);
  }
}

// ===== Instructor Dashboard =====
function renderInstructorDashboard() {
  renderAllStudents();
  renderAllResultsTable();
  populateManualStudentDropdown();
  populateSubjectDropdowns();
  renderSubjectsList();
  renderGradeScaleDisplay();
}

function renderAllStudents() {
  const list = document.getElementById('studentsList');
  const empty = document.getElementById('studentsEmpty');
  const delAllBtn = document.getElementById('deleteAllStudentsBtn');
  const users = Storage.getUsers().filter(u => u.role === 'student');
  list.innerHTML = '';
  if (users.length === 0) {
    empty.classList.remove('hidden');
    if (delAllBtn) delAllBtn.style.display = 'none';
    return;
  }
  empty.classList.add('hidden');
  if (delAllBtn) delAllBtn.style.display = 'inline-block';
  users.forEach(u => {
    const div = document.createElement('div');
    div.className = 'student-item';
    div.innerHTML = `
      <div class="student-info">
        <div>${u.firstName} ${u.fatherName}</div>
        <div class="email">ID: ${u.studentId}</div>
      </div>
      <div style="display:flex;gap:0.4rem;align-items:center;">
        <span class="badge" style="background:#e8edf2;padding:0.2rem 0.6rem;border-radius:12px;font-size:0.75rem;color:#666;">${u.profile.phone || 'No phone'}</span>
        <button class="btn btn-sm btn-danger" onclick="deleteStudent('${u.studentId.replace(/'/g,"\\'")}')">Delete</button>
      </div>
    `;
    list.appendChild(div);
  });
  document.getElementById('studentCount').textContent = users.length;
}

function deleteStudent(sid) {
  if (!confirm(`Delete student ${sid} and all their results?`)) return;
  let users = Storage.getUsers();
  users = users.filter(u => u.studentId !== sid);
  Storage.setUsers(users);
  let results = Storage.getResults();
  results = results.filter(r => r.studentId !== sid);
  Storage.setResults(results);
  renderInstructorDashboard();
}

function deleteAllStudents() {
  if (!confirm('Delete ALL students and their results? This cannot be undone.')) return;
  let users = Storage.getUsers().filter(u => u.role !== 'student');
  Storage.setUsers(users);
  Storage.setResults([]);
  renderInstructorDashboard();
}

function renderAllResultsTable() {
  const tbody = document.getElementById('allResultsBody');
  const empty = document.getElementById('allResultsEmpty');
  const filter = document.getElementById('resultsFilter');
  const selected = filter ? filter.value : '';
  let all = Results.getAll();
  if (selected) all = all.filter(r => r.subject === selected);
  const users = Storage.getUsers().filter(u => u.role === 'student');
  tbody.innerHTML = '';
  if (all.length === 0) {
    empty.classList.remove('hidden');
    return;
  }
  empty.classList.add('hidden');
  // Populate filter dropdown
  if (filter) {
    const current = filter.value;
    filter.innerHTML = '<option value="">All Subjects</option>';
    const subjects = [...new Set(Results.getAll().map(r => r.subject))];
    subjects.sort().forEach(s => {
      const opt = document.createElement('option');
      opt.value = s;
      opt.textContent = s;
      filter.appendChild(opt);
    });
    filter.value = current && subjects.includes(current) ? current : '';
  }
  const grouped = {};
  all.forEach(r => {
    const key = r.studentId;
    if (!grouped[key]) grouped[key] = {};
    if (!grouped[key][r.subject]) grouped[key][r.subject] = { mid: '-', quiz: '-', project: '-', assignment: '-', final: '-', total: 0 };
    const s = grouped[key][r.subject];
    if (r.mid !== null) s.mid = r.mid;
    if (r.quiz !== null) s.quiz = r.quiz;
    if (r.project !== null) s.project = r.project;
    if (r.assignment !== null) s.assignment = r.assignment;
    if (r.final !== null) s.final = r.final;
    s.total = (typeof r.total === 'number' ? r.total : (r.mid || 0) + (r.quiz || 0) + (r.project || 0) + (r.assignment || 0) + (r.final || 0));
  });
  Object.entries(grouped).forEach(([sid, subjects]) => {
    const user = users.find(u => u.studentId === sid);
    const name = user ? `${user.firstName} ${user.fatherName}` : sid;
    Object.entries(subjects).forEach(([subject, scores]) => {
      const grade = calculateGrade(scores.total);
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${name}</td><td>${sid}</td><td>${subject}</td><td>${scores.mid}</td><td>${scores.quiz}</td><td>${scores.project}</td><td>${scores.assignment}</td><td>${scores.final}</td><td><strong>${scores.total}</strong></td><td><strong>${grade.letter}</strong></td>`;
      tbody.appendChild(tr);
    });
  });
}

// Upload
let uploadFile = null;

function handleFileSelect(input) {
  if (input.files && input.files[0]) {
    uploadFile = input.files[0];
    const info = document.getElementById('fileInfo');
    info.classList.remove('hidden');
    info.innerHTML = `<span>📄 ${uploadFile.name}</span> <button class="btn btn-sm btn-danger" onclick="clearFile()">Remove</button>`;
  }
}

function clearFile() {
  uploadFile = null;
  document.getElementById('fileInput').value = '';
  document.getElementById('fileInfo').classList.add('hidden');
}

function handleUpload() {
  const subject = document.getElementById('uploadSubject').value.trim();
  const msgEl = document.getElementById('uploadMsg');
  if (!subject) {
    msgEl.className = 'msg error'; msgEl.textContent = 'Please enter a subject name.'; return;
  }
  if (!uploadFile) {
    msgEl.className = 'msg error'; msgEl.textContent = 'Please select an Excel file.'; return;
  }
  msgEl.className = 'msg';
  msgEl.textContent = 'Uploading...';
  Results.upload(subject, uploadFile, function (result) {
    msgEl.className = result.ok ? 'msg success' : 'msg error';
    msgEl.textContent = result.msg;
    if (result.ok) {
      clearFile();
      renderInstructorDashboard();
    }
  });
}

// Drag and drop
function setupDragDrop() {
  const area = document.getElementById('uploadArea');
  if (!area) return;
  area.addEventListener('dragover', function (e) {
    e.preventDefault();
    this.classList.add('dragover');
  });
  area.addEventListener('dragleave', function () {
    this.classList.remove('dragover');
  });
  area.addEventListener('drop', function (e) {
    e.preventDefault();
    this.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
      uploadFile = e.dataTransfer.files[0];
      const info = document.getElementById('fileInfo');
      info.classList.remove('hidden');
      info.innerHTML = `<span>📄 ${uploadFile.name}</span> <button class="btn btn-sm btn-danger" onclick="clearFile()">Remove</button>`;
    }
  });
  area.addEventListener('click', function () {
    document.getElementById('fileInput').click();
  });
}
document.addEventListener('DOMContentLoaded', setupDragDrop);

// Download sample Excel
function downloadSampleExcel() {
  const wb = XLSX.utils.book_new();
  const data = [
    { 'Student ID': 'STU001', 'Mid Exam': 30, 'Quiz': 10, 'Project': 15, 'Assignment': 10, 'Final Exam': 35, 'Total': 100 },
    { 'Student ID': 'STU002', 'Mid Exam': 25, 'Quiz': 8, 'Project': 14, 'Assignment': 9, 'Final Exam': 30, 'Total': 86 }
  ];
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, 'Results');
  XLSX.writeFile(wb, 'sample_results.xlsx');
}

// Download sample student registration Excel
function downloadSampleStudentExcel() {
  const wb = XLSX.utils.book_new();
  const data = [
    { 'First Name': 'John', 'Father Name': 'Doe', 'Student ID': 'STU001', 'Password': 'student123' },
    { 'First Name': 'Jane', 'Father Name': 'Smith', 'Student ID': 'STU002', 'Password': 'student123' }
  ];
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, 'Students');
  XLSX.writeFile(wb, 'sample_students.xlsx');
}

// Close modals on overlay click
document.addEventListener('click', function (e) {
  if (e.target.classList.contains('modal-overlay')) {
    e.target.classList.add('hidden');
  }
});

// Enter key handlers
document.addEventListener('keydown', function (e) {
  if (e.key === 'Enter') {
    if (!document.getElementById('authSection').classList.contains('hidden')) {
      if (!document.getElementById('loginForm').classList.contains('hidden')) { if (document.activeElement && (document.activeElement.id === 'loginId' || document.activeElement.id === 'loginPassword')) handleLogin(); }
      else if (!document.getElementById('registerForm').classList.contains('hidden')) handleRegister();
      else if (!document.getElementById('resetForm').classList.contains('hidden')) handleReset();
    }
  }
});
