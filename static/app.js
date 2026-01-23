/**
 * LEB Tracker - Frontend Application
 * Vanilla JavaScript for managing RFIs and Submittals
 */

// =============================================================================
// STATE
// =============================================================================

const state = {
    user: null,
    items: [],
    users: [],
    stats: {},
    currentBucket: 'ALL',
    currentTypeFilter: '',
    currentView: 'items', // 'items', 'inbox', or 'workflow'
    closedSectionExpanded: false,  // Whether closed items section is expanded
    selectedItemId: null,
    workflowItems: [],
    selectedItemIds: [],  // For multi-select
    selectMode: false     // Whether we're in selection mode
};

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================

/**
 * Make an API request
 */
async function api(endpoint, options = {}) {
    const defaultOptions = {
        headers: {
            'Content-Type': 'application/json'
        },
        credentials: 'include'
    };
    
    const response = await fetch(`/api${endpoint}`, {
        ...defaultOptions,
        ...options
    });
    
    if (response.status === 401) {
        // Unauthorized - show login
        showLogin();
        throw new Error('Authentication required');
    }
    
    const data = await response.json();
    
    if (!response.ok) {
        throw new Error(data.error || 'API request failed');
    }
    
    return data;
}

/**
 * Format a date string for display
 */
function formatDate(dateStr) {
    if (!dateStr) return '-';
    
    // Handle YYYY-MM-DD format by parsing as local time, not UTC
    // This prevents the "day early" bug when displaying dates
    let date;
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
        // Date-only string - parse as local time
        const [year, month, day] = dateStr.split('-').map(Number);
        date = new Date(year, month - 1, day);
    } else {
        date = new Date(dateStr);
    }
    
    if (isNaN(date.getTime())) return dateStr;
    
    // Format as "Wed, 1/22/26"
    const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const dayName = days[date.getDay()];
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = String(date.getFullYear()).slice(-2);
    
    return `${dayName}, ${month}/${day}/${year}`;
}

/**
 * Format a datetime string for display
 */
function formatDateTime(dateStr) {
    if (!dateStr) return '-';
    
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return dateStr;
    
    return date.toLocaleString('en-US', {
        month: 'short',
        day: 'numeric',
        year: 'numeric',
        hour: 'numeric',
        minute: '2-digit'
    });
}

/**
 * Get initials from a name
 */
function getInitials(name) {
    if (!name) return '?';
    return name.split(' ')
        .map(part => part[0])
        .join('')
        .toUpperCase()
        .slice(0, 2);
}

/**
 * Escape HTML to prevent XSS
 */
function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

/**
 * Check if a date is overdue
 */
function isOverdue(dateStr) {
    if (!dateStr) return false;
    const dueDate = new Date(dateStr);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return dueDate < today;
}

/**
 * Get bucket display name
 */
function getBucketName(bucket) {
    const names = {
        'ALL': 'General',
        'ACC_TURNER': 'ACC Turner',
        'ACC_MORTENSON': 'ACC Mortenson',
        'ACC_FTI': 'ACC FTI'
    };
    return names[bucket] || bucket;
}

// =============================================================================
// UI HELPERS
// =============================================================================

/**
 * Show an element
 */
function show(element) {
    if (typeof element === 'string') {
        element = document.getElementById(element);
    }
    if (element) {
        element.classList.remove('hidden');
    }
}

/**
 * Hide an element
 */
function hide(element) {
    if (typeof element === 'string') {
        element = document.getElementById(element);
    }
    if (element) {
        element.classList.add('hidden');
    }
}

/**
 * Toggle an element
 */
function toggle(element) {
    if (typeof element === 'string') {
        element = document.getElementById(element);
    }
    if (element) {
        element.classList.toggle('hidden');
    }
}

// =============================================================================
// AUTHENTICATION
// =============================================================================

/**
 * Check if user is logged in
 */
async function checkAuth() {
    try {
        state.user = await api('/auth/me');
        showApp();
        await loadInitialData();
    } catch (e) {
        showLogin();
    }
}

/**
 * Show login page
 */
function showLogin() {
    hide('app');
    show('login-page');
    state.user = null;
}

/**
 * Show main app
 */
function showApp() {
    hide('login-page');
    show('app');
    
    // Update user info
    if (state.user) {
        document.getElementById('user-initials').textContent = getInitials(state.user.display_name);
        document.getElementById('user-name').textContent = state.user.display_name;
        document.getElementById('user-email').textContent = state.user.email;
        
        // Show/hide admin features
        const manageUsersBtn = document.getElementById('btn-manage-users');
        const pendingUpdatesNav = document.getElementById('nav-pending-updates');
        
        if (state.user.role !== 'admin') {
            manageUsersBtn.style.display = 'none';
            if (pendingUpdatesNav) pendingUpdatesNav.style.display = 'none';
        } else {
            manageUsersBtn.style.display = 'flex';
            // Show pending updates nav for admins, load count
            if (pendingUpdatesNav) {
                pendingUpdatesNav.style.display = 'flex';
                loadPendingUpdatesCount();
            }
        }
    }
    
    // Load notification count
    updateNotificationCount();
}

/**
 * Handle login form submission
 */
async function handleLogin(e) {
    e.preventDefault();
    
    const email = document.getElementById('login-email').value;
    const password = document.getElementById('login-password').value;
    const errorEl = document.getElementById('login-error');
    
    try {
        errorEl.textContent = '';
        state.user = await api('/auth/login', {
            method: 'POST',
            body: JSON.stringify({ email, password })
        });
        showApp();
        await loadInitialData();
    } catch (e) {
        errorEl.textContent = e.message || 'Login failed';
    }
}

/**
 * Handle logout
 */
async function handleLogout() {
    try {
        await api('/auth/logout', { method: 'POST' });
    } catch (e) {
        // Ignore errors
    }
    showLogin();
}

/**
 * Toggle sidebar collapsed state
 */
function toggleSidebar() {
    const sidebar = document.getElementById('sidebar');
    const btn = document.getElementById('btn-collapse-sidebar');
    sidebar.classList.toggle('collapsed');
    
    if (sidebar.classList.contains('collapsed')) {
        btn.textContent = '▶';
        btn.title = 'Expand Sidebar';
    } else {
        btn.textContent = '◀';
        btn.title = 'Collapse Sidebar';
    }
}

// =============================================================================
// DATA LOADING
// =============================================================================

/**
 * Load initial data
 */
async function loadInitialData() {
    await Promise.all([
        loadItems(),
        loadUsers(),
        loadStats(),
        loadPollingStatus()
    ]);
}

/**
 * Load items list - always load both open and closed items
 */
async function loadItems() {
    try {
        let endpoint = '/items?';
        
        if (state.currentBucket !== 'ALL') {
            endpoint += `bucket=${state.currentBucket}&`;
        }
        
        if (state.currentTypeFilter) {
            endpoint += `type=${state.currentTypeFilter}&`;
        }
        
        // Always load closed items too - we'll separate them in rendering
        endpoint += `show_closed=true&`;
        
        state.items = await api(endpoint);
        renderItems();
    } catch (e) {
        console.error('Failed to load items:', e);
    }
}

/**
 * Load users list
 */
async function loadUsers() {
    try {
        state.users = await api('/users');
        updateAssignedDropdown();
    } catch (e) {
        console.error('Failed to load users:', e);
    }
}

/**
 * Load stats
 */
async function loadStats() {
    try {
        state.stats = await api('/stats');
        updateStats();
        updateTabCounts();
    } catch (e) {
        console.error('Failed to load stats:', e);
    }
}

/**
 * Load polling status
 */
async function loadPollingStatus() {
    try {
        const status = await api('/poll-status');
        updatePollingStatus(status);
    } catch (e) {
        console.error('Failed to load polling status:', e);
    }
}

// =============================================================================
// RENDERING
// =============================================================================

/**
 * Get due date status class for color coding
 */
function getDueDateClass(status) {
    if (!status) return '';
    if (status === 'red') return 'due-date-red';
    if (status === 'yellow') return 'due-date-yellow';
    if (status === 'green') return 'due-date-green';
    return '';
}

/**
 * Generate row HTML for an item
 */
function generateItemRow(item) {
    const typeClass = item.type === 'RFI' ? 'chip-rfi' : 'chip-submittal';
    const priorityClass = item.priority ? `chip-${item.priority.toLowerCase()}` : '';
    const closedClass = item.closed_at ? 'closed-row' : '';
    const insufficientClass = item.is_contractor_window_insufficient ? 'insufficient-warning' : '';
    
    // Check if item has pending contractor update
    const hasPendingUpdate = item.has_pending_update === 1;
    const pendingUpdateClass = hasPendingUpdate ? 'pending-update-row' : '';
    const updateIcon = hasPendingUpdate ? 
        `<span class="update-badge" title="Contractor update pending review${item.update_type === 'content_change' ? ' (Content Changed)' : ' (Due Date Only)'}">🔄</span>` : '';
    
    // Check if item is unread by current user
    const readBy = item.read_by ? item.read_by.split(',').map(id => parseInt(id)) : [];
    const isUnread = state.user && !readBy.includes(state.user.id);
    const unreadDot = isUnread ? '<span class="unread-dot"></span>' : '';
    
    // Warning icon for insufficient window
    const warningIcon = item.is_contractor_window_insufficient ? '<span class="warning-icon-small" title="Insufficient contractor window">⚠️</span>' : '';
    
    // Selection checkbox for select mode
    const isSelected = state.selectedItemIds.includes(item.id);
    const selectCheckbox = state.selectMode ? 
        `<td class="select-cell"><input type="checkbox" class="item-checkbox" ${isSelected ? 'checked' : ''} onclick="event.stopPropagation(); toggleItemSelection(${item.id})"></td>` : '';
    const selectedClass = isSelected ? 'selected-row' : '';
    
    return `
        <tr data-item-id="${item.id}" class="${closedClass} ${insufficientClass} ${selectedClass} ${pendingUpdateClass}">
            ${selectCheckbox}
            <td>${unreadDot}${updateIcon}<span class="chip ${typeClass}">${escapeHtml(item.type)}</span></td>
            <td>${warningIcon}${escapeHtml(item.identifier)}</td>
            <td>${escapeHtml(item.title) || '<span class="text-muted">No title</span>'}</td>
            <td>${formatDate(item.date_received)}</td>
            <td>${item.priority ? `<span class="chip ${priorityClass}">${escapeHtml(item.priority)}</span>` : '-'}</td>
            <td>${formatDate(item.due_date)}</td>
            <td>${formatDate(item.initial_reviewer_due_date)}</td>
            <td>${formatDate(item.qcr_due_date)}</td>
            <td><span class="chip chip-status">${escapeHtml(item.status)}</span></td>
            <td>${escapeHtml(item.initial_reviewer_name) || '<span class="text-muted">-</span>'}</td>
            <td>${escapeHtml(item.qcr_name) || '<span class="text-muted">-</span>'}</td>
        </tr>
    `;
}

/**
 * Render items table
 */
function renderItems() {
    const tbody = document.getElementById('items-table-body');
    const emptyState = document.getElementById('empty-state');
    
    // Separate open and closed items
    const openItems = state.items.filter(item => !item.closed_at);
    const closedItems = state.items.filter(item => item.closed_at);
    
    if (openItems.length === 0 && closedItems.length === 0) {
        tbody.innerHTML = '';
        show(emptyState);
        return;
    }
    
    hide(emptyState);
    
    // Render open items
    let html = openItems.map(item => generateItemRow(item)).join('');
    
    // Add closed items section if there are any
    if (closedItems.length > 0) {
        const isExpanded = state.closedSectionExpanded || false;
        const colSpan = state.selectMode ? 12 : 11;
        
        html += `
            <tr class="closed-section-header" onclick="toggleClosedSection()">
                <td colspan="${colSpan}">
                    <span class="closed-section-toggle">${isExpanded ? '▼' : '▶'}</span>
                    <span class="closed-section-title">Closed Items (${closedItems.length})</span>
                </td>
            </tr>
        `;
        
        if (isExpanded) {
            html += closedItems.map(item => generateItemRow(item)).join('');
        }
    }
    
    tbody.innerHTML = html;
    
    // Update table header for select mode
    const thead = document.querySelector('#items-table thead tr');
    const hasSelectHeader = thead.querySelector('.select-header');
    if (state.selectMode && !hasSelectHeader) {
        thead.insertAdjacentHTML('afterbegin', '<th class="select-header"><input type="checkbox" id="select-all-checkbox" onclick="toggleSelectAll()"></th>');
    } else if (!state.selectMode && hasSelectHeader) {
        hasSelectHeader.remove();
    }
}

/**
 * Toggle closed items section expand/collapse
 */
function toggleClosedSection() {
    state.closedSectionExpanded = !state.closedSectionExpanded;
    renderItems();
}

/**
 * Update stats display
 */
function updateStats() {
    document.getElementById('stat-open').textContent = state.stats.open_items || 0;
    document.getElementById('stat-overdue').textContent = state.stats.overdue_items || 0;
    document.getElementById('stat-due-week').textContent = state.stats.due_this_week || 0;
}

/**
 * Update tab counts
 */
function updateTabCounts() {
    const byBucket = state.stats.by_bucket || {};
    
    // Calculate ALL count (sum of all items)
    const allCount = state.stats.total_items || 0;
    
    document.getElementById('count-all').textContent = allCount;
    document.getElementById('count-turner').textContent = byBucket.ACC_TURNER || 0;
    document.getElementById('count-mortenson').textContent = byBucket.ACC_MORTENSON || 0;
    document.getElementById('count-fti').textContent = byBucket.ACC_FTI || 0;
    
    // Update inbox count
    const inboxCount = state.stats.inbox_count || 0;
    document.getElementById('count-inbox').textContent = inboxCount;
}

/**
 * Update polling status display
 */
function updatePollingStatus(status) {
    const dot = document.getElementById('polling-dot');
    const text = document.getElementById('polling-status');
    
    if (!status.outlook_available) {
        dot.className = 'status-dot error';
        text.textContent = 'Outlook not available';
    } else if (status.running) {
        dot.className = 'status-dot active';
        if (status.last_poll) {
            const lastPoll = new Date(status.last_poll);
            text.textContent = `Last poll: ${lastPoll.toLocaleTimeString()}`;
        } else {
            text.textContent = 'Email polling active';
        }
    } else {
        dot.className = 'status-dot';
        text.textContent = 'Polling inactive';
    }
}

/**
 * Update all reviewer dropdowns - now just updates hidden selects for form data
 * The visible UI uses autocomplete inputs
 */
function updateAssignedDropdown() {
    const userOptions = '<option value="">Not Assigned</option>' +
        state.users.map(user => 
            `<option value="${user.id}">${escapeHtml(user.display_name)}</option>`
        ).join('');
    
    // Update hidden selects (used for form submission)
    const initialReviewerSelect = document.getElementById('detail-initial-reviewer');
    const qcrSelect = document.getElementById('detail-qcr');
    
    if (initialReviewerSelect) initialReviewerSelect.innerHTML = userOptions;
    if (qcrSelect) qcrSelect.innerHTML = userOptions;
    
    // Store users for autocomplete
    allUsersForChips = state.users;
}

// =============================================================================
// DETAIL DRAWER
// =============================================================================

/**
 * Open detail drawer for an item
 */
async function openDetailDrawer(itemId) {
    state.selectedItemId = itemId;
    
    try {
        const item = await api(`/item/${itemId}`);
        populateDetailDrawer(item);
        
        // Load comments
        await loadComments(itemId);
        
        show('drawer-overlay');
        show('detail-drawer');
    } catch (e) {
        console.error('Failed to load item:', e);
    }
}

/**
 * Close detail drawer
 */
function closeDetailDrawer() {
    hide('drawer-overlay');
    hide('detail-drawer');
    state.selectedItemId = null;
}

/**
 * Populate detail drawer with item data
 */
function populateDetailDrawer(item) {
    document.getElementById('drawer-title').textContent = item.identifier;
    
    // Type chip
    const typeEl = document.getElementById('detail-type');
    typeEl.textContent = item.type;
    typeEl.className = `chip ${item.type === 'RFI' ? 'chip-rfi' : 'chip-submittal'}`;
    
    // Info fields
    document.getElementById('detail-bucket').textContent = getBucketName(item.bucket);
    document.getElementById('detail-identifier').textContent = item.identifier;
    document.getElementById('detail-date-received').textContent = formatDate(item.date_received) || '-';
    document.getElementById('detail-created').textContent = formatDateTime(item.created_at);
    document.getElementById('detail-last-email').textContent = formatDateTime(item.last_email_at);
    document.getElementById('detail-subject').textContent = item.source_subject || 'Manual entry';
    
    // Warning banner for insufficient window
    const warningBanner = document.getElementById('insufficient-window-warning');
    const warningText = document.getElementById('warning-text');
    if (item.is_contractor_window_insufficient) {
        warningText.textContent = `⚠️ Contractor provided insufficient review window`;
        show(warningBanner);
    } else {
        hide(warningBanner);
    }
    
    // Contractor update review panel (admin only)
    renderContractorUpdatePanel(item);
    
    // Editable fields
    document.getElementById('detail-title-input').value = item.title || '';
    document.getElementById('detail-due-date').value = item.due_date || '';
    document.getElementById('detail-priority').value = item.priority || '';
    document.getElementById('detail-status').value = item.status || 'Unassigned';
    document.getElementById('detail-folder').value = item.folder_link || '';
    document.getElementById('detail-notes').value = item.notes || '';
    
    // RFI Question field - show only for RFI type
    const rfiQuestionSection = document.getElementById('rfi-question-section');
    const rfiQuestionInput = document.getElementById('detail-rfi-question');
    if (item.type === 'RFI') {
        rfiQuestionInput.value = item.rfi_question || '';
        show(rfiQuestionSection);
    } else {
        rfiQuestionInput.value = '';
        hide(rfiQuestionSection);
    }
    
    // Reviewer fields
    document.getElementById('detail-initial-reviewer').value = item.initial_reviewer_id || '';
    document.getElementById('detail-qcr').value = item.qcr_id || '';
    
    // Calculated due dates - now editable inputs with color coding
    const initialReviewerDue = document.getElementById('detail-initial-reviewer-due');
    const qcrDue = document.getElementById('detail-qcr-due');
    
    initialReviewerDue.value = item.initial_reviewer_due_date || '';
    initialReviewerDue.className = `calculated-date-input ${getDueDateClass(item.initial_reviewer_due_status)}`;
    
    qcrDue.value = item.qcr_due_date || '';
    qcrDue.className = `calculated-date-input ${getDueDateClass(item.qcr_due_status)}`;
    
    // Clear reviewer error
    hide('reviewer-error');
    
    // Check if QCR has provided final response (Approve or Modify action)
    const hasQcrFinalResponse = item.qcr_response_at && (item.qcr_action === 'Approve' || item.qcr_action === 'Modify');
    const responseSection = document.getElementById('response-section');
    
    if (hasQcrFinalResponse) {
        // Hide manual response section when QCR has finalized
        hide(responseSection);
    } else {
        // Show manual response section when no QCR final response yet
        show(responseSection);
    }
    
    // Response fields - use final response values if available, otherwise fall back to item values
    const finalCategory = item.final_response_category || item.response_category || '';
    const finalText = item.final_response_text || item.qcr_notes || item.response_text || '';
    document.getElementById('detail-response-category').value = finalCategory;
    document.getElementById('detail-response-text').value = finalText;
    
    // Load files section - show checkboxes for file selection
    const folderPath = item.folder_link;
    const fileListEl = document.getElementById('file-list');
    
    if (folderPath) {
        loadFileCheckboxes(item.id, item.response_files, folderPath);
    } else {
        fileListEl.innerHTML = '<p class="text-muted">No folder linked. Set a folder path in Files Folder above first.</p>';
    }
    
    // Closeout buttons visibility
    const closeBtn = document.getElementById('btn-close-item');
    const reopenBtn = document.getElementById('btn-reopen-item');
    const closedInfo = document.getElementById('closed-at-info');
    
    if (item.closed_at) {
        hide(closeBtn);
        closedInfo.textContent = `Closed on ${formatDateTime(item.closed_at)}`;
        show(closedInfo);
        
        // Only show reopen button for admins
        if (state.user && state.user.role === 'admin') {
            show(reopenBtn);
        } else {
            hide(reopenBtn);
        }
    } else {
        show(closeBtn);
        hide(reopenBtn);
        hide(closedInfo);
    }
    
    // Update workflow status in drawer
    updateDrawerWorkflowStatus(item);
    
    // Load reviewers using the new chip-based system
    loadReviewerChips(item);
    
    // Initialize QCR autocomplete and set value
    initQcrAutocomplete();
    setQcrFromItem(item);
}

/**
 * Load comments for an item
 */
async function loadComments(itemId) {
    try {
        const comments = await api(`/comments/${itemId}`);
        renderComments(comments);
    } catch (e) {
        console.error('Failed to load comments:', e);
    }
}

/**
 * Render comments list
 */
function renderComments(comments) {
    const list = document.getElementById('comments-list');
    
    if (comments.length === 0) {
        list.innerHTML = '<p class="text-muted">No comments yet</p>';
        return;
    }
    
    list.innerHTML = comments.map(comment => `
        <div class="comment-item">
            <div class="comment-header">
                <span class="comment-author">${escapeHtml(comment.author_name)}</span>
                <span class="comment-date">${formatDateTime(comment.created_at)}</span>
            </div>
            <div class="comment-body">${escapeHtml(comment.body)}</div>
        </div>
    `).join('');
}

/**
 * Validate reviewer selections (Initial Reviewer cannot be the same as QCR)
 */
function validateReviewers() {
    const initialReviewerId = document.getElementById('detail-initial-reviewer').value;
    const qcrId = document.getElementById('detail-qcr').value;
    const errorEl = document.getElementById('reviewer-error');
    
    if (initialReviewerId && qcrId && initialReviewerId === qcrId) {
        errorEl.style.display = 'block';
        return false;
    } else {
        errorEl.style.display = 'none';
        return true;
    }
}

/**
 * Save item changes
 */
async function saveItem() {
    if (!state.selectedItemId) return;
    
    // Validate reviewers first
    if (!validateReviewers()) {
        alert('Initial Reviewer and QCR cannot be the same person.');
        return;
    }
    
    // Get QCR ID - prefer selectedQcrUser, fall back to hidden select
    const qcrId = selectedQcrUser ? selectedQcrUser.id : document.getElementById('detail-qcr').value;
    
    const data = {
        title: document.getElementById('detail-title-input').value,
        due_date: document.getElementById('detail-due-date').value || null,
        priority: document.getElementById('detail-priority').value || null,
        status: document.getElementById('detail-status').value,
        initial_reviewer_id: document.getElementById('detail-initial-reviewer').value || null,
        qcr_id: qcrId || null,
        initial_reviewer_due_date: document.getElementById('detail-initial-reviewer-due').value || null,
        qcr_due_date: document.getElementById('detail-qcr-due').value || null,
        notes: document.getElementById('detail-notes').value,
        rfi_question: document.getElementById('detail-rfi-question').value
    };
    
    try {
        await api(`/item/${state.selectedItemId}`, {
            method: 'POST',
            body: JSON.stringify(data)
        });
        
        // Refresh data
        await loadItems();
        await loadStats();
        
        // Close the drawer
        closeDetailDrawer();
        
    } catch (e) {
        alert('Failed to save: ' + e.message);
    }
}

/**
 * Add a comment
 */
async function addComment() {
    if (!state.selectedItemId) return;
    
    const textarea = document.getElementById('new-comment');
    const body = textarea.value.trim();
    
    if (!body) return;
    
    try {
        await api(`/comments/${state.selectedItemId}`, {
            method: 'POST',
            body: JSON.stringify({ body })
        });
        
        textarea.value = '';
        await loadComments(state.selectedItemId);
    } catch (e) {
        alert('Failed to add comment: ' + e.message);
    }
}

/**
 * Open folder in explorer
 */
function openFolder() {
    const folderPath = document.getElementById('detail-folder').value;
    
    if (!folderPath) {
        alert('No folder path set for this item.');
        return;
    }
    
    openFilesFolder(folderPath);
}

/**
 * Open original email in Outlook
 */
async function openOriginalEmail() {
    console.log('openOriginalEmail called, selectedItemId:', state.selectedItemId);
    
    if (!state.selectedItemId) {
        alert('No item selected.');
        return;
    }
    
    try {
        console.log('Making API call to open email...');
        const result = await api(`/item/${state.selectedItemId}/open-email`, {
            method: 'POST'
        });
        
        console.log('API result:', result);
        if (result.success) {
            console.log('Opened email in Outlook');
        } else {
            alert(result.error || 'Could not open email');
        }
    } catch (e) {
        console.error('Error opening email:', e);
        if (e.message.includes('No original email')) {
            alert('No original email found for this item. This item may have been created manually.');
        } else {
            alert(`Could not open email: ${e.message}`);
        }
    }
}

/**
 * Open a folder path in Windows Explorer via backend API
 */
async function openFilesFolder(folderPath) {
    try {
        const result = await api('/open-folder', {
            method: 'POST',
            body: JSON.stringify({ path: folderPath })
        });
        
        if (result.success) {
            console.log('Opened folder:', result.path);
        }
    } catch (e) {
        // Fallback: copy path to clipboard
        navigator.clipboard.writeText(folderPath).then(() => {
            alert(`Could not open folder automatically.\n\nPath copied to clipboard:\n${folderPath}\n\nPaste into Windows Explorer.`);
        }).catch(() => {
            alert(`Could not open folder.\n\nManually navigate to:\n${folderPath}`);
        });
    }
}

// =============================================================================
// INBOX FUNCTIONS
// =============================================================================

/**
 * Load inbox items (assigned to current user)
 */
async function loadInbox() {
    try {
        state.items = await api('/inbox');
        renderItems();
        
        // Update page title
        document.getElementById('page-title').textContent = 'My Inbox';
    } catch (e) {
        console.error('Failed to load inbox:', e);
    }
}

/**
 * Switch to inbox view
 */
function switchToInbox() {
    state.currentView = 'inbox';
    
    // Update nav active state - remove from all, add to inbox
    document.querySelectorAll('.nav-tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.view === 'inbox');
    });
    
    // Show table container, hide other containers
    show('table-container');
    hide('workflow-container');
    hide('notifications-container');
    
    loadInbox();
}

/**
 * Switch to items view (bucket tabs)
 */
function switchToItems(bucket) {
    state.currentView = 'items';
    
    // Update nav active state
    document.querySelectorAll('.nav-tab').forEach(tab => {
        if (tab.dataset.bucket) {
            tab.classList.toggle('active', tab.dataset.bucket === bucket);
        } else {
            tab.classList.remove('active');
        }
    });
    
    // Show table container, hide other containers
    show('table-container');
    hide('workflow-container');
    hide('notifications-container');
    
    switchBucket(bucket);
}

/**
 * Mark item as read
 */
async function markItemAsRead(itemId) {
    try {
        await api(`/item/${itemId}/mark-read`, { method: 'POST' });
    } catch (e) {
        console.error('Failed to mark as read:', e);
    }
}

// =============================================================================
// WORKFLOW FUNCTIONS
// =============================================================================

/**
 * Load workflow items
 */
async function loadWorkflow() {
    try {
        state.workflowItems = await api('/admin/workflow');
        renderWorkflow();
    } catch (e) {
        console.error('Failed to load workflow:', e);
    }
}

/**
 * Render workflow table
 */
function renderWorkflow() {
    const tbody = document.getElementById('workflow-table-body');
    const emptyState = document.getElementById('workflow-empty-state');
    
    if (!state.workflowItems || state.workflowItems.length === 0) {
        tbody.innerHTML = '';
        show(emptyState);
        return;
    }
    
    hide(emptyState);
    
    tbody.innerHTML = state.workflowItems.map(item => {
        const reviewerStatus = item.reviewer_response_status || 'Not Sent';
        const qcrStatus = item.qcr_response_status || 'Not Sent';
        
        const canSendReviewer = item.reviewer_name && !item.reviewer_email_sent_at;
        const canResendReviewer = item.reviewer_email_sent_at && !item.reviewer_response_at;
        const canResendQcr = item.qcr_email_sent_at && !item.qcr_response_at;
        
        let reviewerActions = '';
        if (canSendReviewer) {
            reviewerActions = `<button class="workflow-action-btn" onclick="sendReviewerEmail(${item.id})">📧 Send</button>`;
        } else if (canResendReviewer) {
            reviewerActions = `<button class="workflow-action-btn" onclick="sendReviewerEmail(${item.id})">🔄 Resend</button>`;
        }
        
        let qcrActions = '';
        if (canResendQcr) {
            qcrActions = `<button class="workflow-action-btn" onclick="sendQcrEmail(${item.id})">🔄 Resend</button>`;
        }
        
        // Format reviewer response details
        let reviewerResponseHtml = '-';
        if (item.reviewer_response_at) {
            const category = item.reviewer_response_category || 'N/A';
            let filesCount = 'No files';
            try {
                const files = item.reviewer_selected_files ? JSON.parse(item.reviewer_selected_files) : [];
                filesCount = files.length > 0 ? `${files.length} file(s)` : 'No files';
            } catch (e) {
                if (item.reviewer_selected_files) filesCount = '✓ Files selected';
            }
            const notes = item.reviewer_notes ? `<div class="response-notes" title="${escapeHtml(item.reviewer_notes)}">${escapeHtml(item.reviewer_notes.substring(0, 50))}${item.reviewer_notes.length > 50 ? '...' : ''}</div>` : '';
            const statusBadge = getReviewerStatusBadge(reviewerStatus);
            reviewerResponseHtml = `
                <div class="response-details">
                    ${statusBadge}
                    <div class="response-category"><strong>${escapeHtml(category)}</strong></div>
                    <div class="response-files">📁 ${escapeHtml(filesCount)}</div>
                    ${notes}
                    <div class="response-time">${formatDateTime(item.reviewer_response_at)}</div>
                </div>
            `;
        } else if (item.reviewer_email_sent_at) {
            reviewerResponseHtml = `<div class="response-pending">⏳ Awaiting response<br><span style="font-size: 0.7rem;">Sent: ${formatDateTime(item.reviewer_email_sent_at)}</span></div>`;
        } else {
            reviewerResponseHtml = `<div class="response-not-sent">⚪ Not sent</div>`;
        }
        
        // Format QC Decision
        let qcDecisionHtml = '-';
        if (item.qcr_response_at && item.qcr_action) {
            const actionIcon = item.qcr_action === 'Approve' ? '✅' : 
                              item.qcr_action === 'Modify' ? '✏️' : '↩️';
            const actionColor = item.qcr_action === 'Approve' ? '#059669' : 
                               item.qcr_action === 'Modify' ? '#2563eb' : '#dc2626';
            const modeText = item.qcr_response_mode ? ` (${item.qcr_response_mode})` : '';
            const qcrNotes = item.qcr_notes ? `<div class="response-notes" title="${escapeHtml(item.qcr_notes)}">${escapeHtml(item.qcr_notes.substring(0, 40))}${item.qcr_notes.length > 40 ? '...' : ''}</div>` : '';
            qcDecisionHtml = `
                <div class="response-details">
                    <div class="qc-action-badge" style="color: ${actionColor}; font-weight: 600;">${actionIcon} ${escapeHtml(item.qcr_action)}</div>
                    ${item.qcr_response_mode ? `<div class="qc-mode" style="font-size: 0.7rem; color: #666;">Mode: ${escapeHtml(item.qcr_response_mode)}</div>` : ''}
                    ${qcrNotes}
                    <div class="response-time">${formatDateTime(item.qcr_response_at)}</div>
                </div>
            `;
        } else if (item.qcr_email_sent_at) {
            qcDecisionHtml = `<div class="response-pending">⏳ Awaiting QC<br><span style="font-size: 0.7rem;">Sent: ${formatDateTime(item.qcr_email_sent_at)}</span></div>`;
        } else if (item.reviewer_response_at) {
            qcDecisionHtml = `<div class="response-pending" style="color: #f59e0b;">🔔 Ready for QC</div>`;
        }
        
        // Format Final Response
        let finalResponseHtml = '-';
        if (item.final_response_category) {
            let finalFilesCount = 'No files';
            try {
                const files = item.final_response_files ? JSON.parse(item.final_response_files) : [];
                finalFilesCount = files.length > 0 ? `${files.length} file(s)` : 'No files';
            } catch (e) {
                if (item.final_response_files) finalFilesCount = '✓ Files';
            }
            finalResponseHtml = `
                <div class="response-details">
                    <div class="response-category" style="color: #059669;"><strong>✅ ${escapeHtml(item.final_response_category)}</strong></div>
                    <div class="response-files">📁 ${escapeHtml(finalFilesCount)}</div>
                </div>
            `;
        } else if (item.qcr_action === 'Send Back') {
            finalResponseHtml = `<div class="response-pending" style="color: #dc2626;">↩️ Sent back for revision</div>`;
        }
        
        // Format folder link
        let folderHtml = '-';
        if (item.folder_link) {
            folderHtml = `<a href="#" onclick="copyFolderPath('${escapeHtml(item.folder_link)}'); return false;" title="Click to copy path" class="folder-link">📂 Copy</a>`;
        }
        
        return `
            <tr onclick="openItem(${item.id})" style="cursor: pointer;">
                <td>
                    <strong>${escapeHtml(item.type)} ${escapeHtml(item.identifier)}</strong>
                    <div class="text-muted" style="font-size: 0.75rem;">${escapeHtml(item.title || '')}</div>
                    <div class="text-muted" style="font-size: 0.7rem; color: #888;">Status: ${escapeHtml(item.status || 'N/A')}</div>
                </td>
                <td>${escapeHtml(item.reviewer_name || '-')}</td>
                <td class="workflow-response-cell">${reviewerResponseHtml}</td>
                <td>${escapeHtml(item.qcr_name || '-')}</td>
                <td class="workflow-response-cell">${qcDecisionHtml}</td>
                <td class="workflow-response-cell">${finalResponseHtml}</td>
                <td>${folderHtml}</td>
                <td onclick="event.stopPropagation();">${reviewerActions}${qcrActions}</td>
            </tr>
        `;
    }).join('');
}

/**
 * Get reviewer status badge HTML
 */
function getReviewerStatusBadge(status) {
    if (status === 'Responded') {
        return '<div class="status-badge responded">✓ Responded</div>';
    } else if (status === 'Email Sent') {
        return '<div class="status-badge email-sent">📧 Sent</div>';
    }
    return '';
}

/**
 * Copy folder path to clipboard
 */
function copyFolderPath(path) {
    navigator.clipboard.writeText(path).then(() => {
        alert('Folder path copied to clipboard!');
    }).catch(err => {
        console.error('Failed to copy path:', err);
        prompt('Copy this path:', path);
    });
}

/**
 * Get CSS class for workflow status badge
 */
function getWorkflowBadgeClass(status) {
    switch (status) {
        case 'Responded':
            return 'responded';
        case 'Email Sent':
            return 'email-sent';
        case 'Not Sent':
        default:
            return 'not-sent';
    }
}

/**
 * Send reviewer email
 */
async function sendReviewerEmail(itemId) {
    try {
        const result = await api(`/admin/send_reviewer_email/${itemId}`, { method: 'POST' });
        alert(result.message || 'Email sent successfully!');
        loadWorkflow();
        // Also reload item details if open
        if (state.selectedItemId === itemId) {
            openItem(itemId);
        }
    } catch (e) {
        alert('Failed to send email: ' + e.message);
    }
}

/**
 * Send QCR email
 */
async function sendQcrEmail(itemId) {
    try {
        const result = await api(`/admin/send_qcr_email/${itemId}`, { method: 'POST' });
        alert(result.message || 'Email sent successfully!');
        loadWorkflow();
    } catch (e) {
        alert('Failed to send email: ' + e.message);
    }
}

/**
 * Handle "Send to Reviewer" button click in drawer
 */
async function handleSendToReviewer() {
    if (!state.selectedItemId) return;
    
    // Capture the item ID before saveItem clears it
    const itemId = state.selectedItemId;
    
    // Check if we have reviewers using chip system
    const hasReviewers = selectedReviewerChips && selectedReviewerChips.length > 0;
    
    // Check QCR - use either the selectedQcrUser object or the hidden select
    const qcrId = selectedQcrUser ? selectedQcrUser.id : document.getElementById('detail-qcr').value;
    
    if (!hasReviewers) {
        alert('Please add at least one Initial Reviewer before sending.');
        return;
    }
    
    if (!qcrId) {
        alert('Please assign a QCR before sending.');
        return;
    }
    
    // Check for conflict - QCR shouldn't be one of the reviewers
    if (selectedQcrUser) {
        const qcrEmail = selectedQcrUser.email.toLowerCase();
        const isConflict = selectedReviewerChips.some(c => c.email.toLowerCase() === qcrEmail);
        if (isConflict) {
            alert('QCR cannot be one of the Initial Reviewers.');
            return;
        }
    }
    
    // Save the item first to ensure QCR is saved
    await saveItem();
    
    // Always use multi-reviewer endpoint since reviewers are stored in item_reviewers table
    await handleSendToReviewers(itemId, qcrId);
}

/**
 * Switch to workflow view
 */
function switchToWorkflow() {
    state.currentView = 'workflow';
    
    // Update nav active state
    document.querySelectorAll('.nav-tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.view === 'workflow');
    });
    
    // Show workflow container, hide table container
    hide('table-container');
    hide('notifications-container');
    hide('pending-updates-container');
    show('workflow-container');
    
    // Update header
    document.getElementById('page-title').textContent = 'Email & Responses';
    document.getElementById('page-subtitle').textContent = 'Reviewer Workflow Status';
    
    loadWorkflow();
}

/**
 * Switch to pending updates view (admin only)
 */
function switchToPendingUpdates() {
    state.currentView = 'pending-updates';
    
    // Update nav active state
    document.querySelectorAll('.nav-tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.view === 'pending-updates');
    });
    
    // Show pending updates container, hide other containers
    hide('table-container');
    hide('workflow-container');
    hide('notifications-container');
    hide('reminders-container');
    show('pending-updates-container');
    
    // Update header
    document.getElementById('page-title').textContent = 'Contractor Updates';
    document.getElementById('page-subtitle').textContent = 'Items Updated by Contractor';
    
    loadPendingUpdates();
}

/**
 * Load pending contractor updates
 */
async function loadPendingUpdates() {
    try {
        const updates = await api('/pending-updates');
        renderPendingUpdates(updates);
    } catch (err) {
        console.error('Failed to load pending updates:', err);
        showToast('Failed to load pending updates', 'error');
    }
}

/**
 * Render pending contractor updates list
 */
function renderPendingUpdates(updates) {
    const container = document.getElementById('pending-updates-list');
    const emptyState = document.getElementById('pending-updates-empty-state');
    
    if (!updates || updates.length === 0) {
        container.innerHTML = '';
        show('pending-updates-empty-state');
        return;
    }
    
    hide('pending-updates-empty-state');
    
    container.innerHTML = updates.map(item => {
        const isContentChange = item.update_type === 'content_change';
        const wasReopened = item.reopened_from_closed === 1;
        const typeClass = item.type === 'RFI' ? 'chip-rfi' : 'chip-submittal';
        
        let changesText = [];
        if (item.previous_due_date) changesText.push('Due Date');
        if (item.previous_title) changesText.push('Title');
        if (item.previous_priority) changesText.push('Priority');
        
        return `
            <div class="pending-update-card ${isContentChange ? 'content-change' : 'due-date-only'}">
                <div class="update-card-header">
                    <div class="update-card-title">
                        <span class="chip ${typeClass}">${escapeHtml(item.type)}</span>
                        <strong>${escapeHtml(item.identifier)}</strong>
                        ${wasReopened ? '<span class="reopened-badge">REOPENED</span>' : ''}
                    </div>
                    <div class="update-card-type">
                        ${isContentChange ? '⚠️ Content Change' : '📅 Due Date Change'}
                    </div>
                </div>
                <div class="update-card-body">
                    <p class="update-card-title-text">${escapeHtml(item.title || 'No title')}</p>
                    <p class="update-card-changes">Changes: ${changesText.join(', ') || 'Unknown'}</p>
                    <p class="update-card-meta">
                        <span>Detected: ${formatDateTime(item.update_detected_at)}</span>
                        ${item.status_before_update ? `<span>Previous Status: ${item.status_before_update}</span>` : ''}
                    </p>
                </div>
                <div class="update-card-actions">
                    <button class="btn btn-primary btn-sm" onclick="openItem(${item.id})">
                        Review Update
                    </button>
                </div>
            </div>
        `;
    }).join('');
}

/**
 * Switch to notifications view
 */
function switchToNotifications() {
    state.currentView = 'notifications';
    
    // Update nav active state
    document.querySelectorAll('.nav-tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.view === 'notifications');
    });
    
    // Show notifications container, hide other containers
    hide('table-container');
    hide('workflow-container');
    hide('pending-updates-container');
    show('notifications-container');
    
    // Update header
    document.getElementById('page-title').textContent = 'Notifications';
    document.getElementById('page-subtitle').textContent = 'System Updates & Action Items';
    
    loadNotifications();
}

/**
 * Load notifications from API
 */
async function loadNotifications() {
    try {
        const response = await api('/notifications');
        renderNotifications(response.notifications);
    } catch (err) {
        console.error('Failed to load notifications:', err);
        showToast('Failed to load notifications', 'error');
    }
}

/**
 * Render notifications list
 */
function renderNotifications(notifications) {
    const container = document.getElementById('notifications-list');
    const emptyState = document.getElementById('notifications-empty-state');
    
    if (!notifications || notifications.length === 0) {
        container.innerHTML = '';
        show('notifications-empty-state');
        return;
    }
    
    hide('notifications-empty-state');
    
    container.innerHTML = notifications.map(n => {
        const isUnread = !n.read_at;
        const icon = getNotificationIcon(n.type);
        const timeAgo = formatTimeAgo(n.created_at);
        
        let actionsHtml = '';
        // Add View Item button if item_id is present
        if (n.item_id) {
            actionsHtml += `<button class="btn btn-outline btn-sm" onclick="viewNotificationItem(${n.item_id}, ${n.id})">👁️ View Item</button>`;
        }
        // For response_ready with action_url, use markItemComplete instead of handleNotificationAction
        if (n.type === 'response_ready' && n.item_id) {
            actionsHtml = `<button class="btn btn-success btn-sm" onclick="markItemComplete(${n.item_id}, ${n.id})">✓ Mark Complete</button>` + actionsHtml;
        } else if (n.action_url && n.action_label) {
            actionsHtml = `<button class="btn btn-primary btn-sm" onclick="handleNotificationAction(${n.id}, '${n.action_url}')">${n.action_label}</button>` + actionsHtml;
        }
        if (isUnread) {
            actionsHtml += `<button class="btn btn-secondary btn-sm" onclick="markNotificationRead(${n.id})">Mark Read</button>`;
        }
        actionsHtml += `<button class="btn btn-secondary btn-sm" onclick="deleteNotification(${n.id})">×</button>`;
        
        return `
            <div class="notification-item type-${n.type} ${isUnread ? 'unread' : ''}" data-id="${n.id}">
                <div class="notification-icon">${icon}</div>
                <div class="notification-content">
                    <div class="notification-title">${escapeHtml(n.title)}</div>
                    <div class="notification-message">${escapeHtml(n.message)}</div>
                    <div class="notification-meta">
                        <span class="notification-time">🕐 ${timeAgo}</span>
                    </div>
                </div>
                <div class="notification-actions">
                    ${actionsHtml}
                </div>
            </div>
        `;
    }).join('');
}

/**
 * Get icon for notification type
 */
function getNotificationIcon(type) {
    const icons = {
        'response_ready': '✅',
        'sent_back': '↩️',
        'info': 'ℹ️',
        'warning': '⚠️',
        'error': '❌'
    };
    return icons[type] || '🔔';
}

// =============================================================================
// REMINDERS VIEW
// =============================================================================

/**
 * Switch to reminders view
 */
function switchToReminders() {
    state.currentView = 'reminders';
    
    // Update nav active state
    document.querySelectorAll('.nav-tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.view === 'reminders');
    });
    
    // Show reminders container, hide other containers
    hide('table-container');
    hide('workflow-container');
    hide('notifications-container');
    show('reminders-container');
    
    // Update header
    document.getElementById('page-title').textContent = 'Reminders';
    document.getElementById('page-subtitle').textContent = 'Reminder Management';
    
    loadRemindersView();
}

/**
 * Load reminders view data
 */
async function loadRemindersView() {
    await Promise.all([
        loadReminderStatus(),
        loadPendingReminders(),
        loadReminderHistory()
    ]);
}

/**
 * Load reminder scheduler status
 */
async function loadReminderStatus() {
    try {
        const status = await api('/reminder-status');
        document.getElementById('reminder-scheduler-status').textContent = status.running ? '🟢 Running' : '🔴 Stopped';
        document.getElementById('reminder-scheduler-status').style.color = status.running ? '#059669' : '#dc2626';
        document.getElementById('reminder-time').textContent = status.reminder_time_pst;
        document.getElementById('reminder-last-check').textContent = status.last_check ? formatDateTime(status.last_check) : 'Never';
    } catch (err) {
        console.error('Failed to load reminder status:', err);
    }
}

/**
 * Load pending reminders
 */
async function loadPendingReminders() {
    try {
        const data = await api('/pending-reminders');
        renderPendingReminders(data);
    } catch (err) {
        console.error('Failed to load pending reminders:', err);
        document.getElementById('pending-reminders-list').innerHTML = '<p class="text-muted">Failed to load pending reminders</p>';
    }
}

/**
 * Render pending reminders
 */
function renderPendingReminders(data) {
    const container = document.getElementById('pending-reminders-list');
    
    const allPending = [
        ...data.single_reviewer.map(r => ({ ...r, mode: 'single', displayName: r.recipient })),
        ...data.multi_reviewer.map(r => ({ ...r, mode: 'multi', displayName: r.reviewer_name })),
        ...data.multi_reviewer_qcr.map(r => ({ ...r, mode: 'multi_qcr', displayName: r.qcr_email, role: 'qcr' }))
    ];
    
    if (allPending.length === 0) {
        container.innerHTML = '<div class="empty-reminder-state"><p>✅ No pending reminders at this time</p></div>';
        return;
    }
    
    container.innerHTML = allPending.map(r => {
        const stageClass = r.reminder_stage === 'overdue' ? 'overdue' : 'due-today';
        const stageLabel = r.reminder_stage === 'overdue' ? '⚠️ OVERDUE' : '⏰ Due Today';
        const roleLabel = r.role === 'qcr' ? 'QCR' : 'Reviewer';
        
        return `
            <div class="pending-reminder-item ${stageClass}">
                <div class="reminder-item-info">
                    <div class="reminder-identifier">${r.identifier}</div>
                    <div class="reminder-recipient">${roleLabel}: ${r.displayName}</div>
                    <div class="reminder-due">Due: ${r.due_date}</div>
                </div>
                <div class="reminder-item-status">
                    <span class="reminder-stage ${stageClass}">${stageLabel}</span>
                </div>
                <div class="reminder-item-actions">
                    <button class="btn btn-primary btn-sm" onclick="sendItemReminder(${r.item_id})">📧 Send Reminder</button>
                </div>
            </div>
        `;
    }).join('');
}

/**
 * Load reminder history
 */
async function loadReminderHistory() {
    try {
        const data = await api('/reminder-history');
        renderReminderHistory(data.reminders);
    } catch (err) {
        console.error('Failed to load reminder history:', err);
        document.getElementById('reminder-history-list').innerHTML = '<p class="text-muted">Failed to load reminder history</p>';
    }
}

/**
 * Render reminder history
 */
function renderReminderHistory(reminders) {
    const container = document.getElementById('reminder-history-list');
    
    if (!reminders || reminders.length === 0) {
        container.innerHTML = '<div class="empty-reminder-state"><p>No reminders have been sent yet</p></div>';
        return;
    }
    
    container.innerHTML = reminders.map(r => {
        const stageClass = r.reminder_stage === 'overdue' ? 'overdue' : r.reminder_stage === 'due_today' ? 'due-today' : 'manual';
        const stageLabel = r.reminder_stage === 'overdue' ? '⚠️ Overdue' : r.reminder_stage === 'due_today' ? '⏰ Due Today' : '✋ Manual';
        const roleLabel = r.role === 'qcr' ? 'QCR' : 'Reviewer';
        
        return `
            <div class="reminder-history-item">
                <div class="history-item-info">
                    <div class="history-identifier">${r.identifier || `Item #${r.item_id}`}</div>
                    <div class="history-recipient">${roleLabel}: ${r.recipient_email}</div>
                    <div class="history-sent">Sent: ${formatDateTime(r.sent_at)}</div>
                </div>
                <div class="history-item-meta">
                    <span class="reminder-stage ${stageClass}">${stageLabel}</span>
                    <span class="reminder-mode">${r.reminder_type}</span>
                </div>
            </div>
        `;
    }).join('');
}

/**
 * Send reminder for a specific item
 */
async function sendItemReminder(itemId) {
    if (!confirm('Send reminder for this item?')) return;
    
    try {
        const result = await api(`/items/${itemId}/send-reminder`, { method: 'POST' });
        if (result.success) {
            const successCount = result.results.filter(r => r.success).length;
            showToast(`Sent ${successCount} reminder(s)`, 'success');
            if (state.currentView === 'reminders') {
                loadRemindersView();
            }
        } else {
            showToast(result.error || 'Failed to send reminder', 'error');
        }
    } catch (err) {
        console.error('Failed to send reminder:', err);
        showToast('Failed to send reminder', 'error');
    }
}

/**
 * Handle send reminder button in drawer
 */
async function handleSendReminderFromDrawer() {
    if (!state.selectedItemId) {
        showToast('No item selected', 'error');
        return;
    }
    await sendItemReminder(state.selectedItemId);
}

/**
 * Process all pending reminders
 */
async function processAllReminders() {
    if (!confirm('Send all pending reminders now?')) return;
    
    try {
        const result = await api('/process-reminders', { method: 'POST' });
        if (result.success) {
            showToast('Reminders processed successfully', 'success');
            loadRemindersView();
        } else {
            showToast(result.error || 'Failed to process reminders', 'error');
        }
    } catch (err) {
        console.error('Failed to process reminders:', err);
        showToast('Failed to process reminders', 'error');
    }
}

/**
 * Format time ago string
 */
function formatTimeAgo(dateStr) {
    if (!dateStr) return '';
    const date = new Date(dateStr);
    const now = new Date();
    const diffMs = now - date;
    const diffMins = Math.floor(diffMs / 60000);
    const diffHours = Math.floor(diffMs / 3600000);
    const diffDays = Math.floor(diffMs / 86400000);
    
    if (diffMins < 1) return 'Just now';
    if (diffMins < 60) return `${diffMins} min ago`;
    if (diffHours < 24) return `${diffHours} hour${diffHours > 1 ? 's' : ''} ago`;
    if (diffDays < 7) return `${diffDays} day${diffDays > 1 ? 's' : ''} ago`;
    return formatDate(dateStr);
}

/**
 * Mark a notification as read
 */
async function markNotificationRead(id) {
    try {
        await api(`/notifications/${id}/read`, { method: 'POST' });
        loadNotifications();
        updateNotificationCount();
    } catch (err) {
        console.error('Failed to mark notification read:', err);
    }
}

/**
 * Mark all notifications as read
 */
async function markAllNotificationsRead() {
    try {
        await api('/notifications/read-all', { method: 'POST' });
        loadNotifications();
        updateNotificationCount();
        showToast('All notifications marked as read', 'success');
    } catch (err) {
        console.error('Failed to mark all notifications read:', err);
        showToast('Failed to mark notifications read', 'error');
    }
}

/**
 * Delete a notification
 */
async function deleteNotification(id) {
    try {
        await api(`/notifications/${id}`, { method: 'DELETE' });
        await loadNotifications();
        await updateNotificationCount();
    } catch (err) {
        console.error('Failed to delete notification:', err);
        showToast('Failed to delete notification', 'error');
    }
}

/**
 * Handle notification action button click
 */
async function handleNotificationAction(notificationId, actionUrl) {
    // Mark as read first
    await markNotificationRead(notificationId);
    
    // If it's an API URL, make a POST call; otherwise navigate
    if (actionUrl.startsWith('/api/')) {
        try {
            await api(actionUrl.replace('/api', ''), { method: 'POST' });
            showToast('Action completed successfully', 'success');
            loadNotifications();
            loadItems();
        } catch (err) {
            console.error('Failed to perform action:', err);
            showToast('Failed to perform action', 'error');
        }
    } else if (actionUrl.startsWith('/')) {
        window.location.href = actionUrl;
    }
}

/**
 * Mark an item as complete
 */
async function markItemComplete(itemId, notificationId) {
    try {
        await api(`/items/${itemId}/complete`, { method: 'POST' });
        showToast('Item marked as complete', 'success');
        
        // Delete the notification since it's been acted upon
        if (notificationId) {
            await deleteNotification(notificationId);
        }
        
        // Refresh relevant views
        if (state.currentView === 'notifications') {
            loadNotifications();
        }
        loadItems();
    } catch (err) {
        console.error('Failed to mark item complete:', err);
        showToast('Failed to mark item complete', 'error');
    }
}

/**
 * View item from notification - opens the item drawer
 */
async function viewNotificationItem(itemId, notificationId) {
    // Mark notification as read
    if (notificationId) {
        await markNotificationRead(notificationId);
    }
    
    // Close notifications panel
    hide('notifications-container');
    
    // Open item drawer
    openDetailDrawer(itemId);
}

/**
 * Update notification count badge
 */
async function updateNotificationCount() {
    try {
        const response = await api('/notifications');
        const unreadCount = response.notifications.filter(n => !n.read_at).length;
        const badge = document.getElementById('count-notifications');
        if (badge) {
            badge.textContent = unreadCount > 0 ? unreadCount : '';
            badge.dataset.count = unreadCount;
        }
    } catch (err) {
        console.error('Failed to update notification count:', err);
    }
}

/**
 * Update send button state based on current selections
 */
function updateSendButtonState() {
    const sendBtn = document.getElementById('btn-send-to-reviewer');
    if (!sendBtn) return;
    
    // Check if we have reviewers using the chip system
    const hasReviewers = selectedReviewerChips && selectedReviewerChips.length > 0;
    
    // Check if QCR is selected
    const qcrId = document.getElementById('detail-qcr').value;
    const hasQcr = qcrId && qcrId !== '';
    
    // Check for same reviewer (QCR shouldn't be in reviewer chips)
    const qcrEmail = selectedQcrUser ? selectedQcrUser.email.toLowerCase() : '';
    const sameReviewer = hasReviewers && selectedReviewerChips.some(c => c.email.toLowerCase() === qcrEmail);
    
    if (!hasReviewers) {
        sendBtn.disabled = true;
        sendBtn.textContent = '📧 Add Reviewer First';
    } else if (!hasQcr) {
        sendBtn.disabled = true;
        sendBtn.textContent = '📧 Assign QCR First';
    } else if (sameReviewer) {
        sendBtn.disabled = true;
        sendBtn.textContent = '⚠️ QCR Cannot Be Reviewer';
    } else {
        sendBtn.disabled = false;
        const reviewerCount = selectedReviewerChips.length;
        sendBtn.textContent = reviewerCount > 1 
            ? `📧 Send to ${reviewerCount} Reviewers` 
            : '📧 Send to Reviewer';
    }
}

/**
 * Update workflow status display in drawer
 */
function updateDrawerWorkflowStatus(item) {
    const workflowActions = document.getElementById('workflow-actions');
    const statusInfo = document.getElementById('workflow-status-info');
    const statusText = document.getElementById('workflow-status-text');
    const sendBtn = document.getElementById('btn-send-to-reviewer');
    
    // First update based on current dropdown values
    updateSendButtonState();
    
    // Then check workflow status from item data to override button state if needed
    const emailSent = item.reviewer_email_sent_at;
    const reviewerResponded = item.reviewer_response_at;
    const qcrResponded = item.qcr_response_at;
    
    // Override button state based on workflow progress
    if (emailSent && reviewerResponded && qcrResponded) {
        sendBtn.disabled = true;
        sendBtn.textContent = '✅ Review Complete';
    } else if (emailSent) {
        sendBtn.disabled = false;
        sendBtn.textContent = '🔄 Resend to Reviewer';
    }
    
    // Show status info if workflow has started
    if (emailSent) {
        let statusHtml = '';
        
        statusHtml += `<p><strong>Reviewer Email:</strong> Sent ${formatDateTime(item.reviewer_email_sent_at)}</p>`;
        
        if (reviewerResponded) {
            statusHtml += `<p class="success"><strong>Reviewer Response:</strong> ${formatDateTime(item.reviewer_response_at)}</p>`;
            statusHtml += `<p><strong>Reviewer Category:</strong> ${item.reviewer_response_category || 'N/A'}</p>`;
            if (item.reviewer_notes) {
                statusHtml += `<p style="margin: 8px 0 3px 0;"><strong>Reviewer Description:</strong></p>`;
                statusHtml += `<div style="background: #f9fafb; padding: 8px; border-radius: 4px; margin-top: 4px; white-space: pre-wrap; font-size: 0.9rem; word-wrap: break-word;">${escapeHtml(item.reviewer_notes)}</div>`;
            }
            // Show reviewer internal notes (team only)
            if (item.reviewer_internal_notes) {
                statusHtml += `<div style="background: #fff8e6; padding: 8px; border-radius: 4px; margin-top: 8px; border: 1px solid #ffd966;">`;
                statusHtml += `<p style="margin: 0 0 5px 0;"><strong style="color: #b7791f;">🔒 Reviewer's Internal Notes (Team Only):</strong></p>`;
                statusHtml += `<div style="color: #744210; white-space: pre-wrap; font-size: 0.9rem;">${escapeHtml(item.reviewer_internal_notes)}</div>`;
                statusHtml += `</div>`;
            }
            // Show reviewer selected files
            if (item.reviewer_selected_files) {
                try {
                    const files = JSON.parse(item.reviewer_selected_files);
                    if (files.length > 0) {
                        statusHtml += `<p><strong>Reviewer Files:</strong> ${files.length} file(s) selected</p>`;
                    }
                } catch(e) {}
            }
        } else {
            statusHtml += `<p class="info"><strong>Reviewer Status:</strong> Awaiting Response</p>`;
        }
        
        if (item.qcr_email_sent_at) {
            statusHtml += `<hr style="margin: 10px 0; border-color: #e5e7eb;">`;
            statusHtml += `<p><strong>QCR Email:</strong> Sent ${formatDateTime(item.qcr_email_sent_at)}</p>`;
            
            if (qcrResponded) {
                // Show QCR action with color coding
                const actionIcon = item.qcr_action === 'Approve' ? '✅' : item.qcr_action === 'Modify' ? '✏️' : '↩️';
                const actionColor = item.qcr_action === 'Approve' ? '#059669' : item.qcr_action === 'Modify' ? '#2563eb' : '#dc2626';
                statusHtml += `<p class="success"><strong>QCR Response:</strong> ${formatDateTime(item.qcr_response_at)}</p>`;
                statusHtml += `<p><strong>QCR Action:</strong> <span style="color: ${actionColor}; font-weight: bold;">${actionIcon} ${item.qcr_action || 'N/A'}</span></p>`;
                if (item.qcr_response_mode) {
                    statusHtml += `<p><strong>Response Mode:</strong> ${item.qcr_response_mode}</p>`;
                }
                if (item.qcr_notes) {
                    statusHtml += `<p style="margin: 8px 0 3px 0;"><strong>QCR Notes:</strong></p>`;
                    statusHtml += `<div style="background: #f9fafb; padding: 8px; border-radius: 4px; margin-top: 4px; white-space: pre-wrap; font-size: 0.9rem; word-wrap: break-word;">${escapeHtml(item.qcr_notes)}</div>`;
                }
                // Show QCR internal notes (team only)
                if (item.qcr_internal_notes) {
                    statusHtml += `<div style="background: #fff8e6; padding: 8px; border-radius: 4px; margin-top: 8px; border: 1px solid #ffd966;">`;
                    statusHtml += `<p style="margin: 0 0 5px 0;"><strong style="color: #b7791f;">🔒 QCR's Internal Notes (Team Only):</strong></p>`;
                    statusHtml += `<div style="color: #744210; white-space: pre-wrap; font-size: 0.9rem;">${escapeHtml(item.qcr_internal_notes)}</div>`;
                    statusHtml += `</div>`;
                }
                
                // Show QCR's modified response text if they edited it (Tweak/Revise mode)
                if (item.qcr_response_text && item.qcr_response_mode && item.qcr_response_mode !== 'Keep') {
                    statusHtml += `<div style="background: #eff6ff; padding: 8px; border-radius: 4px; margin-top: 8px; border-left: 3px solid #2563eb;">`;
                    statusHtml += `<p style="margin: 0 0 5px 0;"><strong style="color: #2563eb;">✏️ QCR Modified Response:</strong></p>`;
                    statusHtml += `<p style="margin: 0; white-space: pre-wrap;">${escapeHtml(item.qcr_response_text)}</p>`;
                    statusHtml += `</div>`;
                }
                
                // Show final response if available
                if (item.final_response_category) {
                    statusHtml += `<hr style="margin: 10px 0; border-color: #10b981;">`;
                    statusHtml += `<div style="background: #ecfdf5; padding: 10px; border-radius: 6px; margin-top: 8px; position: relative;">`;
                    statusHtml += `<button onclick="showManualModifyResponse()" style="position: absolute; top: 8px; right: 8px; background: #f3f4f6; border: 1px solid #d1d5db; border-radius: 4px; padding: 4px 8px; cursor: pointer; font-size: 12px;" title="Manually modify response">✏️</button>`;
                    statusHtml += `<p style="margin: 0 0 5px 0;"><strong style="color: #059669;">📋 FINAL RESPONSE</strong></p>`;
                    statusHtml += `<p style="margin: 3px 0;"><strong>Category:</strong> ${item.final_response_category}</p>`;
                    // Include QCR notes in final response
                    if (item.qcr_notes) {
                        statusHtml += `<p style="margin: 8px 0 3px 0;"><strong>Response Notes:</strong></p>`;
                        statusHtml += `<div style="background: white; padding: 8px; border-radius: 4px; margin-top: 4px; white-space: pre-wrap; font-size: 0.9rem;">${escapeHtml(item.qcr_notes)}</div>`;
                    }
                    if (item.final_response_text && item.final_response_text !== item.qcr_notes) {
                        statusHtml += `<p style="margin: 8px 0 3px 0;"><strong>Response Text:</strong></p>`;
                        statusHtml += `<div style="background: white; padding: 8px; border-radius: 4px; margin-top: 4px; white-space: pre-wrap; font-size: 0.9rem;">${escapeHtml(item.final_response_text)}</div>`;
                    }
                    if (item.final_response_files) {
                        try {
                            const finalFiles = JSON.parse(item.final_response_files);
                            if (finalFiles.length > 0) {
                                statusHtml += `<p style="margin: 8px 0 3px 0;"><strong>Files:</strong></p><ul style="margin: 3px 0; padding-left: 20px;">`;
                                finalFiles.forEach(f => {
                                    statusHtml += `<li style="font-size: 0.85rem;">${escapeHtml(f)}</li>`;
                                });
                                statusHtml += `</ul>`;
                            }
                        } catch(e) {}
                    }
                    statusHtml += `</div>`;
                }
            } else {
                statusHtml += `<p class="info"><strong>QCR Status:</strong> Awaiting Response</p>`;
            }
        }
        
        statusText.innerHTML = statusHtml;
        show(statusInfo);
    } else {
        hide(statusInfo);
    }
}

// =============================================================================
// RESPONSE FUNCTIONS
// =============================================================================

/**
 * Load file checkboxes for an item
 */
async function loadFileCheckboxes(itemId, selectedFiles, folderPath) {
    const fileListEl = document.getElementById('file-list');
    
    // Show loading state with open folder button
    fileListEl.innerHTML = `
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.5rem;">
            <span class="text-muted">Loading files...</span>
            <button class="btn btn-outline btn-sm" id="btn-open-files-folder">📂 Open Folder</button>
        </div>
    `;
    document.getElementById('btn-open-files-folder').addEventListener('click', () => openFilesFolder(folderPath));
    
    try {
        const data = await api(`/item/${itemId}/files`);
        const selectedArray = selectedFiles ? selectedFiles.split(',').map(f => f.trim()) : [];
        
        let html = `
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.5rem;">
                <span class="text-muted" style="font-size: 0.75rem;">${data.files.length} file(s) found</span>
                <button class="btn btn-outline btn-sm" id="btn-open-files-folder">📂 Open Folder</button>
            </div>
        `;
        
        if (data.files.length === 0) {
            html += '<p class="text-muted">No files in folder yet. Click "Open Folder" to add files.</p>';
        } else {
            html += '<div class="file-checkbox-items">';
            data.files.forEach((file, index) => {
                const isChecked = selectedArray.includes(file.filename) ? 'checked' : '';
                const fileId = `file-${index}`;
                html += `
                    <div class="file-checkbox-item">
                        <input type="checkbox" id="${fileId}" value="${escapeHtml(file.filename)}" ${isChecked}>
                        <label for="${fileId}">${escapeHtml(file.filename)}</label>
                    </div>
                `;
            });
            html += '</div>';
        }
        
        fileListEl.innerHTML = html;
        
        // Re-attach open folder button listener
        document.getElementById('btn-open-files-folder').addEventListener('click', () => openFilesFolder(folderPath));
        
    } catch (e) {
        fileListEl.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.5rem;">
                <span class="text-muted" style="color: var(--color-danger);">Error loading files</span>
                <button class="btn btn-outline btn-sm" id="btn-open-files-folder">📂 Open Folder</button>
            </div>
            <p class="text-muted">${escapeHtml(e.message)}</p>
        `;
        document.getElementById('btn-open-files-folder').addEventListener('click', () => openFilesFolder(folderPath));
    }
}

/**
 * Save response
 */
async function saveResponse() {
    if (!state.selectedItemId) return;
    
    // Get selected files from checkboxes
    const fileCheckboxes = document.querySelectorAll('#file-list input[type="checkbox"]:checked');
    const selectedFiles = Array.from(fileCheckboxes).map(cb => cb.value);
    
    const data = {
        response_category: document.getElementById('detail-response-category').value || null,
        response_text: document.getElementById('detail-response-text').value || null,
        response_files: selectedFiles.length > 0 ? selectedFiles : null
    };
    
    try {
        await api(`/item/${state.selectedItemId}/response`, {
            method: 'POST',
            body: JSON.stringify(data)
        });
        
        // Show success
        const btn = document.getElementById('btn-save-response');
        const originalText = btn.textContent;
        btn.textContent = '✓ Saved!';
        setTimeout(() => btn.textContent = originalText, 1500);
        
        // Reload items to refresh the display
        await loadItems();
        
        // Re-select the item to update the drawer
        if (state.selectedItemId) {
            const updatedItem = state.items.find(i => i.id === state.selectedItemId);
            if (updatedItem) {
                populateDetailDrawer(updatedItem);
            }
        }
    } catch (e) {
        alert('Failed to save response: ' + e.message);
    }
}

/**
 * Show the manual modify response section (when QCR has already provided final response)
 * This allows the user to manually adjust the response before closeout
 */
function showManualModifyResponse() {
    const responseSection = document.getElementById('response-section');
    
    // Show response section for editing (it's already prefilled from populateDetailDrawer)
    show(responseSection);
    
    // Scroll to response section
    responseSection.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

// =============================================================================
// CLOSEOUT FUNCTIONS
// =============================================================================

/**
 * Close out an item
 */
async function closeItem() {
    if (!state.selectedItemId) return;
    
    if (!confirm('Are you sure you want to close out this item? This will mark it as completed.')) {
        return;
    }
    
    try {
        const result = await api(`/item/${state.selectedItemId}/close`, { method: 'POST' });
        
        // Update drawer UI
        hide('btn-close-item');
        document.getElementById('closed-at-info').textContent = `Closed on ${formatDateTime(result.closed_at)}`;
        show('closed-at-info');
        
        if (state.user && state.user.role === 'admin') {
            show('btn-reopen-item');
        }
        
        // Refresh items
        await loadItems();
        await loadStats();
    } catch (e) {
        alert('Failed to close item: ' + e.message);
    }
}

/**
 * Reopen a closed item (admin only)
 */
async function reopenItem() {
    if (!state.selectedItemId) return;
    
    if (!confirm('Are you sure you want to reopen this item?')) {
        return;
    }
    
    try {
        await api(`/item/${state.selectedItemId}/reopen`, { method: 'POST' });
        
        // Update drawer UI
        show('btn-close-item');
        hide('btn-reopen-item');
        hide('closed-at-info');
        
        // Refresh items
        await loadItems();
        await loadStats();
    } catch (e) {
        alert('Failed to reopen item: ' + e.message);
    }
}

/**
 * Delete an item (admin only) - single item from drawer
 */
async function deleteItem() {
    if (!state.selectedItemId) return;
    
    const deleteFolder = confirm('Do you also want to DELETE the folder and all its files?\n\nClick OK to delete folder, Cancel to keep folder.');
    
    if (!confirm(`ARE YOU SURE you want to permanently delete this item?\n\nThis action CANNOT be undone.`)) {
        return;
    }
    
    try {
        await api(`/items/${state.selectedItemId}?delete_folder=${deleteFolder}`, { method: 'DELETE' });
        
        // Close drawer first
        closeDetailDrawer();
        
        // Remove item from state immediately for instant UI feedback
        state.items = state.items.filter(item => item.id !== state.selectedItemId);
        renderItems();
        
        // Then reload from server to ensure sync
        await loadItems();
        await loadStats();
        
        alert('Item deleted successfully');
    } catch (e) {
        alert('Failed to delete item: ' + e.message);
    }
}

/**
 * Toggle selection mode for multi-select
 */
function toggleSelectMode() {
    state.selectMode = !state.selectMode;
    state.selectedItemIds = [];
    
    const btn = document.getElementById('btn-select-mode');
    const deleteBtn = document.getElementById('btn-delete-selected');
    const cancelBtn = document.getElementById('btn-cancel-select');
    
    if (state.selectMode) {
        btn.textContent = '☑️ Selection Mode ON';
        btn.classList.add('active');
        show('btn-delete-selected');
        show('btn-cancel-select');
        document.getElementById('items-table').classList.add('select-mode');
    } else {
        btn.textContent = '☐ Select Items';
        btn.classList.remove('active');
        hide('btn-delete-selected');
        hide('btn-cancel-select');
        document.getElementById('items-table').classList.remove('select-mode');
    }
    
    updateSelectionCount();
    renderItems();
}

/**
 * Cancel selection mode
 */
function cancelSelectMode() {
    state.selectMode = false;
    state.selectedItemIds = [];
    
    const btn = document.getElementById('btn-select-mode');
    btn.textContent = '☐ Select Items';
    btn.classList.remove('active');
    hide('btn-delete-selected');
    hide('btn-cancel-select');
    document.getElementById('items-table').classList.remove('select-mode');
    
    updateSelectionCount();
    renderItems();
}

/**
 * Toggle item selection
 */
function toggleItemSelection(itemId) {
    const idx = state.selectedItemIds.indexOf(itemId);
    if (idx === -1) {
        state.selectedItemIds.push(itemId);
    } else {
        state.selectedItemIds.splice(idx, 1);
    }
    updateSelectionCount();
    renderItems();
}

/**
 * Update selection count display
 */
function updateSelectionCount() {
    const deleteBtn = document.getElementById('btn-delete-selected');
    if (state.selectedItemIds.length > 0) {
        deleteBtn.textContent = `🗑️ Delete Selected (${state.selectedItemIds.length})`;
        deleteBtn.disabled = false;
    } else {
        deleteBtn.textContent = '🗑️ Delete Selected (0)';
        deleteBtn.disabled = true;
    }
}

/**
 * Delete all selected items
 */
async function deleteSelectedItems() {
    if (state.selectedItemIds.length === 0) {
        alert('No items selected');
        return;
    }
    
    const count = state.selectedItemIds.length;
    const deleteFolder = confirm(`Do you want to DELETE the folders for all ${count} items?\n\nClick OK to delete folders, Cancel to keep folders.`);
    
    if (!confirm(`ARE YOU SURE you want to permanently delete ${count} item(s)?\n\nThis action CANNOT be undone.`)) {
        return;
    }
    
    let successCount = 0;
    let errorCount = 0;
    
    for (const itemId of state.selectedItemIds) {
        try {
            await api(`/items/${itemId}?delete_folder=${deleteFolder}`, { method: 'DELETE' });
            successCount++;
        } catch (e) {
            console.error(`Failed to delete item ${itemId}:`, e);
            errorCount++;
        }
    }
    
    // Exit selection mode first
    cancelSelectMode();
    
    // Refresh the data
    await loadItems();
    await loadStats();
    
    // Show result message using toast instead of blocking alert
    if (errorCount > 0) {
        showToast(`Deleted ${successCount} item(s). ${errorCount} failed.`, 'warning');
    } else {
        showToast(`Successfully deleted ${successCount} item(s).`, 'success');
    }
}

/**
 * Toggle select all items
 */
function toggleSelectAll() {
    const checkbox = document.getElementById('select-all-checkbox');
    if (checkbox.checked) {
        // Select all visible items
        state.selectedItemIds = state.items.map(item => item.id);
    } else {
        // Deselect all
        state.selectedItemIds = [];
    }
    updateSelectionCount();
    renderItems();
}

// =============================================================================
// NEW ITEM MODAL
// =============================================================================

/**
 * Open new item modal
 */
function openNewItemModal() {
    show('new-item-modal');
}

/**
 * Close new item modal
 */
function closeNewItemModal() {
    hide('new-item-modal');
    document.getElementById('new-item-form').reset();
}

/**
 * Handle new item form submission
 */
async function handleNewItem(e) {
    e.preventDefault();
    
    const data = {
        type: document.getElementById('new-type').value,
        bucket: document.getElementById('new-bucket').value,
        identifier: document.getElementById('new-identifier').value,
        title: document.getElementById('new-title').value || null,
        due_date: document.getElementById('new-due-date').value || null,
        priority: document.getElementById('new-priority').value || null
    };
    
    try {
        await api('/items', {
            method: 'POST',
            body: JSON.stringify(data)
        });
        
        closeNewItemModal();
        await loadItems();
        await loadStats();
    } catch (e) {
        alert('Failed to create item: ' + e.message);
    }
}

// =============================================================================
// ACC MODAL (PLACEHOLDER)
// =============================================================================

function openAccModal() {
    show('acc-modal');
}

function closeAccModal() {
    hide('acc-modal');
}

// =============================================================================
// USERS MODAL
// =============================================================================

function openUsersModal() {
    renderUsersList();
    show('users-modal');
}

function closeUsersModal() {
    hide('users-modal');
}

function renderUsersList() {
    const list = document.getElementById('users-list');
    
    list.innerHTML = state.users.map(user => `
        <div class="user-row">
            <div class="user-avatar">
                <span>${getInitials(user.display_name)}</span>
            </div>
            <div class="user-info">
                <div class="user-info-name">${escapeHtml(user.display_name)}</div>
                <div class="user-info-email">${escapeHtml(user.email)}</div>
            </div>
            <span class="user-role">${user.role}</span>
        </div>
    `).join('');
}

// =============================================================================
// OUTLOOK CONTACT AUTOCOMPLETE
// =============================================================================

let autocompleteTimeout = null;

async function searchOutlookContacts(query) {
    if (!query || query.length < 2) {
        hideAutocomplete();
        return;
    }
    
    try {
        const results = await api(`/outlook/contacts?q=${encodeURIComponent(query)}`);
        showAutocomplete(results);
    } catch (e) {
        console.error('Outlook search error:', e);
        hideAutocomplete();
    }
}

function showAutocomplete(contacts) {
    const dropdown = document.getElementById('email-autocomplete');
    
    if (!contacts || contacts.length === 0) {
        hideAutocomplete();
        return;
    }
    
    dropdown.innerHTML = contacts.map(contact => `
        <div class="autocomplete-item" data-email="${escapeHtml(contact.email)}" data-name="${escapeHtml(contact.display_name)}">
            <div class="autocomplete-name">${escapeHtml(contact.display_name)}</div>
            <div class="autocomplete-email">${escapeHtml(contact.email)}</div>
        </div>
    `).join('');
    
    show(dropdown);
    
    // Add click handlers to items
    dropdown.querySelectorAll('.autocomplete-item').forEach(item => {
        item.addEventListener('click', () => selectAutocompleteItem(item));
    });
}

function hideAutocomplete() {
    hide(document.getElementById('email-autocomplete'));
}

function selectAutocompleteItem(item) {
    const email = item.dataset.email;
    const name = item.dataset.name;
    
    document.getElementById('new-user-email').value = email;
    
    // Auto-fill display name if empty
    const nameInput = document.getElementById('new-user-name');
    if (!nameInput.value.trim()) {
        nameInput.value = name;
    }
    
    hideAutocomplete();
}

function setupEmailAutocomplete() {
    const emailInput = document.getElementById('new-user-email');
    
    emailInput.addEventListener('input', (e) => {
        // Debounce the search
        if (autocompleteTimeout) {
            clearTimeout(autocompleteTimeout);
        }
        
        autocompleteTimeout = setTimeout(() => {
            searchOutlookContacts(e.target.value);
        }, 300);  // Wait 300ms after typing stops
    });
    
    emailInput.addEventListener('focus', (e) => {
        if (e.target.value.length >= 2) {
            searchOutlookContacts(e.target.value);
        }
    });
    
    emailInput.addEventListener('blur', (e) => {
        // Delay hiding to allow click on dropdown items
        setTimeout(hideAutocomplete, 200);
    });
    
    // Keyboard navigation
    emailInput.addEventListener('keydown', (e) => {
        const dropdown = document.getElementById('email-autocomplete');
        if (dropdown.classList.contains('hidden')) return;
        
        const items = dropdown.querySelectorAll('.autocomplete-item');
        const activeItem = dropdown.querySelector('.autocomplete-item.active');
        
        if (e.key === 'ArrowDown') {
            e.preventDefault();
            if (!activeItem) {
                items[0]?.classList.add('active');
            } else {
                const idx = Array.from(items).indexOf(activeItem);
                activeItem.classList.remove('active');
                items[(idx + 1) % items.length]?.classList.add('active');
            }
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            if (!activeItem) {
                items[items.length - 1]?.classList.add('active');
            } else {
                const idx = Array.from(items).indexOf(activeItem);
                activeItem.classList.remove('active');
                items[(idx - 1 + items.length) % items.length]?.classList.add('active');
            }
        } else if (e.key === 'Enter' && activeItem) {
            e.preventDefault();
            selectAutocompleteItem(activeItem);
        } else if (e.key === 'Escape') {
            hideAutocomplete();
        }
    });
}

// =============================================================================
// USER MANAGEMENT
// =============================================================================

async function handleNewUser(e) {
    e.preventDefault();
    
    const email = document.getElementById('new-user-email').value.trim();
    const displayName = document.getElementById('new-user-name').value.trim();
    const password = document.getElementById('new-user-password').value;
    const role = document.getElementById('new-user-role').value;
    
    if (!email) {
        alert('Email is required');
        return;
    }
    
    const data = {
        email: email,
        display_name: displayName || null,
        role: role
    };
    
    // Only include password if provided
    if (password) {
        data.password = password;
    }
    
    try {
        await api('/users', {
            method: 'POST',
            body: JSON.stringify(data)
        });
        
        document.getElementById('new-user-form').reset();
        hideAutocomplete();
        await loadUsers();
        renderUsersList();
    } catch (e) {
        alert('Failed to create user: ' + e.message);
    }
}

// =============================================================================
// SETTINGS MODAL
// =============================================================================

async function openSettingsModal() {
    try {
        const config = await api('/config');
        document.getElementById('setting-project-name').value = config.project_name || '';
        document.getElementById('setting-base-folder').value = config.base_folder_path || '';
        document.getElementById('setting-outlook-folder').value = config.outlook_folder || '';
        document.getElementById('setting-poll-interval').value = config.poll_interval_minutes || 5;
        show('settings-modal');
    } catch (e) {
        alert('Failed to load settings: ' + e.message);
    }
}

function closeSettingsModal() {
    hide('settings-modal');
}

async function handleSaveSettings(e) {
    e.preventDefault();
    
    const data = {
        project_name: document.getElementById('setting-project-name').value,
        base_folder_path: document.getElementById('setting-base-folder').value,
        outlook_folder: document.getElementById('setting-outlook-folder').value,
        poll_interval_minutes: parseInt(document.getElementById('setting-poll-interval').value) || 5
    };
    
    try {
        await api('/config', {
            method: 'POST',
            body: JSON.stringify(data)
        });
        
        closeSettingsModal();
        
        // Update project name in UI
        document.getElementById('project-name').textContent = data.project_name;
    } catch (e) {
        alert('Failed to save settings: ' + e.message);
    }
}

// =============================================================================
// NAVIGATION
// =============================================================================

/**
 * Switch to a bucket tab
 */
function switchBucket(bucket) {
    state.currentBucket = bucket;
    
    // Update active tab
    document.querySelectorAll('.nav-tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.bucket === bucket);
    });
    
    // Update header
    const titles = {
        'ALL': 'All Items',
        'ACC_TURNER': 'ACC Turner',
        'ACC_MORTENSON': 'ACC Mortenson',
        'ACC_FTI': 'ACC FTI'
    };
    document.getElementById('page-title').textContent = titles[bucket] || bucket;
    
    // Reload items
    loadItems();
}

/**
 * Set type filter
 */
function setTypeFilter(type) {
    state.currentTypeFilter = type;
    
    // Update active filter chip
    document.querySelectorAll('.filter-chip').forEach(chip => {
        chip.classList.toggle('active', chip.dataset.value === type);
    });
    
    // Reload items
    loadItems();
}

// =============================================================================
// EVENT LISTENERS
// =============================================================================

document.addEventListener('DOMContentLoaded', function() {
    // Login form
    document.getElementById('login-form').addEventListener('submit', handleLogin);
    
    // Navigation tabs - handle both bucket tabs and inbox/workflow tabs
    document.querySelectorAll('.nav-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            if (tab.dataset.view === 'inbox') {
                switchToInbox();
            } else if (tab.dataset.view === 'workflow') {
                switchToWorkflow();
            } else if (tab.dataset.view === 'notifications') {
                switchToNotifications();
            } else if (tab.dataset.view === 'reminders') {
                switchToReminders();
            } else if (tab.dataset.view === 'pending-updates') {
                switchToPendingUpdates();
            } else if (tab.dataset.bucket) {
                switchToItems(tab.dataset.bucket);
            }
        });
    });
    
    // Refresh pending updates button
    const refreshUpdatesBtn = document.getElementById('btn-refresh-updates');
    if (refreshUpdatesBtn) {
        refreshUpdatesBtn.addEventListener('click', loadPendingUpdates);
    }
    
    // Mark All Read button
    const markAllReadBtn = document.getElementById('btn-mark-all-read');
    if (markAllReadBtn) {
        markAllReadBtn.addEventListener('click', markAllNotificationsRead);
    }
    
    // Type filter chips
    document.querySelectorAll('.filter-chip').forEach(chip => {
        chip.addEventListener('click', () => setTypeFilter(chip.dataset.value));
    });
    
    // Sidebar collapse toggle
    document.getElementById('btn-collapse-sidebar').addEventListener('click', toggleSidebar);
    
    // Items table row clicks
    document.getElementById('items-table-body').addEventListener('click', (e) => {
        // Ignore clicks on checkboxes
        if (e.target.type === 'checkbox') return;
        
        const row = e.target.closest('tr');
        if (row && row.dataset.itemId) {
            const itemId = parseInt(row.dataset.itemId);
            
            // In select mode, toggle selection instead of opening drawer
            if (state.selectMode) {
                toggleItemSelection(itemId);
            } else {
                openDetailDrawer(itemId);
                markItemAsRead(itemId);
            }
        }
    });
    
    // Selection mode buttons
    document.getElementById('btn-select-mode').addEventListener('click', toggleSelectMode);
    document.getElementById('btn-delete-selected').addEventListener('click', deleteSelectedItems);
    document.getElementById('btn-cancel-select').addEventListener('click', cancelSelectMode);
    
    // Detail drawer
    document.getElementById('drawer-overlay').addEventListener('click', closeDetailDrawer);
    document.getElementById('btn-close-drawer').addEventListener('click', closeDetailDrawer);
    document.getElementById('btn-save-item').addEventListener('click', saveItem);
    document.getElementById('btn-add-comment').addEventListener('click', addComment);
    document.getElementById('btn-open-folder').addEventListener('click', openFolder);
    document.getElementById('btn-open-email').addEventListener('click', openOriginalEmail);
    
    // QCR dropdown change handler
    document.getElementById('detail-qcr').addEventListener('change', () => {
        validateReviewers();
        updateSendButtonState();
    });
    
    // Response and closeout buttons
    document.getElementById('btn-save-response').addEventListener('click', saveResponse);
    document.getElementById('btn-close-item').addEventListener('click', closeItem);
    document.getElementById('btn-reopen-item').addEventListener('click', reopenItem);
    document.getElementById('btn-delete-item').addEventListener('click', deleteItem);
    
    // Send to Reviewer button
    document.getElementById('btn-send-to-reviewer').addEventListener('click', handleSendToReviewer);
    
    // Send Reminder button in drawer
    document.getElementById('btn-send-reminder').addEventListener('click', handleSendReminderFromDrawer);
    
    // New item modal
    document.getElementById('btn-new-item').addEventListener('click', openNewItemModal);
    document.getElementById('btn-close-new-item').addEventListener('click', closeNewItemModal);
    document.getElementById('new-item-form').addEventListener('submit', handleNewItem);
    
    // ACC modal
    document.getElementById('btn-connect-acc').addEventListener('click', openAccModal);
    document.getElementById('btn-import-acc').addEventListener('click', openAccModal);
    document.getElementById('btn-close-acc').addEventListener('click', closeAccModal);
    
    // Users modal
    document.getElementById('btn-manage-users').addEventListener('click', openUsersModal);
    document.getElementById('btn-close-users').addEventListener('click', closeUsersModal);
    document.getElementById('new-user-form').addEventListener('submit', handleNewUser);
    setupEmailAutocomplete();  // Setup Outlook contact search
    
    // Settings modal
    document.getElementById('btn-settings').addEventListener('click', openSettingsModal);
    document.getElementById('btn-close-settings').addEventListener('click', closeSettingsModal);
    document.getElementById('settings-form').addEventListener('submit', handleSaveSettings);
    
    // Reminder buttons
    const processRemindersBtn = document.getElementById('btn-process-reminders');
    if (processRemindersBtn) {
        processRemindersBtn.addEventListener('click', processAllReminders);
    }
    const refreshRemindersBtn = document.getElementById('btn-refresh-reminders');
    if (refreshRemindersBtn) {
        refreshRemindersBtn.addEventListener('click', loadRemindersView);
    }
    
    // User menu
    document.getElementById('user-avatar').addEventListener('click', () => {
        toggle('user-dropdown');
    });
    
    document.getElementById('btn-logout').addEventListener('click', handleLogout);
    
    // Close dropdown when clicking outside
    document.addEventListener('click', (e) => {
        const dropdown = document.getElementById('user-dropdown');
        const avatar = document.getElementById('user-avatar');
        if (!dropdown.contains(e.target) && !avatar.contains(e.target)) {
            hide(dropdown);
        }
    });
    
    // Close modals on overlay click
    document.querySelectorAll('.modal-overlay').forEach(overlay => {
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) {
                hide(overlay);
            }
        });
    });
    
    // Keyboard shortcuts
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            closeDetailDrawer();
            closeNewItemModal();
            closeAccModal();
            closeUsersModal();
            closeSettingsModal();
            hide('user-dropdown');
        }
    });
    
    // Periodic refresh
    setInterval(async () => {
        if (state.user) {
            await loadStats();
            await loadPollingStatus();
            await updateNotificationCount();
        }
    }, 60000); // Every minute
    
    // Initial auth check
    checkAuth();
});

// =============================================================================
// REVIEWER CHIPS FUNCTIONS (Tag-based multi-reviewer selection)
// =============================================================================

/**
 * State for reviewer chips
 */
let selectedReviewerChips = [];
let allUsersForChips = [];

/**
 * Load reviewer chips for an item - shows existing reviewers as chips and allows adding more
 */
async function loadReviewerChips(item) {
    const container = document.getElementById('reviewer-chips-container');
    const statusList = document.getElementById('reviewer-status-list');
    const searchInput = document.getElementById('reviewer-search-input');
    const dropdown = document.getElementById('reviewer-dropdown');
    
    if (!container) return;
    
    // Load all users for dropdown
    try {
        allUsersForChips = await api('/users');
    } catch (e) {
        console.error('Failed to load users for chips:', e);
        allUsersForChips = [];
    }
    
    // Load existing reviewers for this item
    let existingReviewers = [];
    if (item.id) {
        try {
            existingReviewers = await api(`/item/${item.id}/reviewers`);
        } catch (e) {
            console.error('Failed to load existing reviewers:', e);
        }
    }
    
    // Initialize selected chips from existing reviewers
    selectedReviewerChips = existingReviewers.map(r => ({
        id: r.id,
        user_id: r.reviewer_email, // Use email as identifier
        name: r.reviewer_name,
        email: r.reviewer_email,
        status: r.status || 'pending',
        response_category: r.response_category,
        isExisting: true
    }));
    
    // Render chips
    renderReviewerChips();
    
    // Render status list for existing reviewers with status
    renderReviewerStatusList(existingReviewers);
    
    // Set up search input
    if (searchInput) {
        searchInput.value = '';
        searchInput.oninput = () => filterReviewerDropdown(searchInput.value);
        searchInput.onfocus = () => {
            filterReviewerDropdown(searchInput.value);
            dropdown.classList.add('show');
        };
    }
    
    // Close dropdown when clicking outside
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.reviewer-chips-section')) {
            dropdown?.classList.remove('show');
        }
    });
}

/**
 * Render the reviewer chips in the input area
 */
function renderReviewerChips() {
    const chipsArea = document.getElementById('reviewer-chips-area');
    if (!chipsArea) return;
    
    chipsArea.innerHTML = selectedReviewerChips.map((chip, index) => {
        const statusClass = chip.status === 'responded' ? 'responded' : 
                           chip.status === 'sent' ? 'sent' : '';
        return `
            <span class="reviewer-chip ${statusClass}" data-index="${index}">
                <span class="chip-name">${escapeHtml(chip.name)}</span>
                <button class="chip-remove" onclick="removeReviewerChip(${index})" title="Remove ${escapeHtml(chip.name)}">×</button>
            </span>
        `;
    }).join('');
    
    // Update the Send button text based on reviewer count
    updateSendToReviewerButton();
}

/**
 * Filter and show the reviewer dropdown based on search text
 */
function filterReviewerDropdown(searchText) {
    const dropdown = document.getElementById('reviewer-dropdown');
    if (!dropdown) return;
    
    const search = searchText.toLowerCase();
    
    // Filter users that aren't already selected
    const selectedEmails = selectedReviewerChips.map(c => c.email.toLowerCase());
    const available = allUsersForChips.filter(user => {
        const isSelected = selectedEmails.includes(user.email.toLowerCase());
        const matchesSearch = !search || 
            user.display_name.toLowerCase().includes(search) || 
            user.email.toLowerCase().includes(search);
        return !isSelected && matchesSearch;
    });
    
    if (available.length === 0) {
        dropdown.innerHTML = '<div class="dropdown-empty">No matching users</div>';
    } else {
        dropdown.innerHTML = available.map(user => `
            <div class="dropdown-item" onclick="addReviewerChip('${escapeHtml(user.display_name)}', '${escapeHtml(user.email)}')">
                <span class="dropdown-name">${escapeHtml(user.display_name)}</span>
                <span class="dropdown-email">${escapeHtml(user.email)}</span>
            </div>
        `).join('');
    }
    
    dropdown.classList.add('show');
}

/**
 * Add a reviewer chip
 */
async function addReviewerChip(name, email) {
    // Check if already selected
    if (selectedReviewerChips.some(c => c.email.toLowerCase() === email.toLowerCase())) {
        return;
    }
    
    const chip = {
        name: name,
        email: email,
        status: 'pending',
        isExisting: false
    };
    
    // If we have an item selected, save to database immediately
    if (state.selectedItemId) {
        try {
            const result = await api(`/item/${state.selectedItemId}/reviewers`, {
                method: 'POST',
                body: JSON.stringify({
                    reviewer_name: name,
                    reviewer_email: email
                })
            });
            chip.id = result.reviewer_id;  // API returns reviewer_id, not id
            chip.isExisting = true;
        } catch (e) {
            console.error('Failed to add reviewer:', e);
            alert('Failed to add reviewer: ' + e.message);
            return;
        }
    }
    
    selectedReviewerChips.push(chip);
    renderReviewerChips();
    
    // Clear and hide dropdown
    const searchInput = document.getElementById('reviewer-search-input');
    const dropdown = document.getElementById('reviewer-dropdown');
    if (searchInput) searchInput.value = '';
    if (dropdown) dropdown.classList.remove('show');
    
    // Update button state
    updateSendButtonState();
    
    // Reload status list
    if (state.selectedItemId) {
        const reviewers = await api(`/item/${state.selectedItemId}/reviewers`);
        renderReviewerStatusList(reviewers);
    }
}

/**
 * Remove a reviewer chip
 */
async function removeReviewerChip(index) {
    const chip = selectedReviewerChips[index];
    if (!chip) return;
    
    // If it's saved in database, delete it
    if (chip.isExisting && chip.id && state.selectedItemId) {
        try {
            await api(`/item/${state.selectedItemId}/reviewers/${chip.id}`, {
                method: 'DELETE'
            });
        } catch (e) {
            console.error('Failed to remove reviewer:', e);
            alert('Failed to remove reviewer: ' + e.message);
            return;
        }
    }
    
    selectedReviewerChips.splice(index, 1);
    renderReviewerChips();
    
    // Update button state
    updateSendButtonState();
    
    // Reload status list
    if (state.selectedItemId) {
        const reviewers = await api(`/item/${state.selectedItemId}/reviewers`);
        renderReviewerStatusList(reviewers);
    }
}

/**
 * Render the status list showing reviewer response status
 */
function renderReviewerStatusList(reviewers) {
    const statusList = document.getElementById('reviewer-status-list');
    if (!statusList) return;
    
    const sentOrResponded = reviewers.filter(r => r.status && r.status !== 'pending');
    
    if (sentOrResponded.length === 0) {
        statusList.innerHTML = '';
        return;
    }
    
    statusList.innerHTML = `
        <div class="status-list-header">Reviewer Status:</div>
        ${sentOrResponded.map(r => {
            const statusClass = r.status === 'responded' ? 'responded' : 'sent';
            const statusText = r.status === 'responded' ? 
                `✓ ${r.response_category || 'Responded'}` : 
                '📧 Email Sent';
            return `
                <div class="reviewer-status-item ${statusClass}">
                    <span class="status-name">${escapeHtml(r.reviewer_name)}</span>
                    <span class="status-badge">${statusText}</span>
                </div>
            `;
        }).join('')}
    `;
}

/**
 * Update the Send to Reviewer button based on selected chips
 */
function updateSendToReviewerButton() {
    // Just call the main update function which handles all logic
    updateSendButtonState();
}

/**
 * Handle sending to reviewer(s) - automatically uses multi-reviewer if more than one selected
 */
async function handleSendToReviewers(itemId = null, qcrIdParam = null) {
    // Use passed itemId or fall back to state
    const targetItemId = itemId || state.selectedItemId;
    if (!targetItemId) return;
    
    const reviewerCount = selectedReviewerChips.length;
    
    // Get QCR ID from param, selectedQcrUser, or hidden select
    const qcrId = qcrIdParam || (selectedQcrUser ? selectedQcrUser.id : document.getElementById('detail-qcr').value);
    
    if (reviewerCount === 0) {
        alert('Please select at least one reviewer by typing their name above.');
        return;
    }
    
    if (!qcrId) {
        alert('Please assign a QCR before sending to reviewers.');
        return;
    }
    
    try {
        // Save QCR assignment and set multi-reviewer mode if needed
        await api(`/item/${targetItemId}`, {
            method: 'POST',
            body: JSON.stringify({
                qcr_id: qcrId,
                multi_reviewer_mode: reviewerCount > 1 ? 1 : 0
            })
        });
        
        // Send emails to all reviewers
        const result = await api(`/item/${targetItemId}/send-multi-reviewer-emails`, {
            method: 'POST'
        });
        
        alert(result.message || 'Email(s) sent successfully!');
        
        // Reload chips and items
        const reviewers = await api(`/item/${targetItemId}/reviewers`);
        selectedReviewerChips = reviewers.map(r => ({
            id: r.id,
            name: r.reviewer_name,
            email: r.reviewer_email,
            status: r.status || 'pending',
            response_category: r.response_category,
            isExisting: true
        }));
        renderReviewerChips();
        renderReviewerStatusList(reviewers);
        
        await loadItems();
    } catch (e) {
        alert('Failed to send emails: ' + e.message);
    }
}

// =============================================================================
// QCR AUTOCOMPLETE FUNCTIONS
// =============================================================================

let selectedQcrUser = null;

/**
 * Initialize QCR autocomplete
 */
function initQcrAutocomplete() {
    const searchInput = document.getElementById('qcr-search-input');
    const dropdown = document.getElementById('qcr-dropdown');
    
    if (!searchInput) return;
    
    searchInput.oninput = () => filterQcrDropdown(searchInput.value);
    searchInput.onfocus = () => {
        filterQcrDropdown(searchInput.value);
        dropdown.classList.add('show');
    };
    
    // Close dropdown when clicking outside
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.qcr-autocomplete-container')) {
            dropdown?.classList.remove('show');
        }
    });
}

/**
 * Filter and show the QCR dropdown based on search text
 */
function filterQcrDropdown(searchText) {
    const dropdown = document.getElementById('qcr-dropdown');
    if (!dropdown) return;
    
    const search = searchText.toLowerCase();
    
    // Filter users - exclude currently selected reviewers to avoid conflict
    const selectedReviewerEmails = selectedReviewerChips.map(c => c.email.toLowerCase());
    const available = allUsersForChips.filter(user => {
        const matchesSearch = !search || 
            user.display_name.toLowerCase().includes(search) || 
            user.email.toLowerCase().includes(search);
        return matchesSearch;
    });
    
    if (available.length === 0) {
        dropdown.innerHTML = '<div class="dropdown-empty">No matching users</div>';
    } else {
        dropdown.innerHTML = available.map(user => {
            const isSelected = selectedQcrUser && selectedQcrUser.id === user.id;
            const isReviewer = selectedReviewerEmails.includes(user.email.toLowerCase());
            return `
                <div class="dropdown-item ${isSelected ? 'selected' : ''} ${isReviewer ? 'reviewer-warning' : ''}" 
                     onclick="selectQcr(${user.id}, '${escapeHtml(user.display_name)}', '${escapeHtml(user.email)}')"
                     ${isReviewer ? 'title="This user is already a reviewer"' : ''}>
                    <span class="dropdown-name">${escapeHtml(user.display_name)}${isReviewer ? ' ⚠️' : ''}</span>
                    <span class="dropdown-email">${escapeHtml(user.email)}</span>
                </div>
            `;
        }).join('');
    }
    
    dropdown.classList.add('show');
}

/**
 * Select a QCR user
 */
function selectQcr(userId, name, email) {
    const searchInput = document.getElementById('qcr-search-input');
    const dropdown = document.getElementById('qcr-dropdown');
    const hiddenSelect = document.getElementById('detail-qcr');
    
    // Update hidden select value
    if (hiddenSelect) {
        hiddenSelect.value = userId;
    }
    
    // Update search input display
    if (searchInput) {
        searchInput.value = name;
        searchInput.classList.add('has-value');
    }
    
    // Store selected user
    selectedQcrUser = { id: userId, name: name, email: email };
    
    // Hide dropdown
    if (dropdown) {
        dropdown.classList.remove('show');
    }
    
    // Validate and update button state
    validateReviewers();
    updateSendButtonState();
}

/**
 * Clear QCR selection
 */
function clearQcrSelection() {
    const searchInput = document.getElementById('qcr-search-input');
    const hiddenSelect = document.getElementById('detail-qcr');
    
    if (searchInput) {
        searchInput.value = '';
        searchInput.classList.remove('has-value');
    }
    
    if (hiddenSelect) {
        hiddenSelect.value = '';
    }
    
    selectedQcrUser = null;
}

/**
 * Set QCR from item data (when loading an item)
 */
function setQcrFromItem(item) {
    const searchInput = document.getElementById('qcr-search-input');
    const hiddenSelect = document.getElementById('detail-qcr');
    
    if (item.qcr_id && item.qcr_name) {
        if (searchInput) {
            searchInput.value = item.qcr_name;
            searchInput.classList.add('has-value');
        }
        if (hiddenSelect) {
            hiddenSelect.value = item.qcr_id;
        }
        selectedQcrUser = { id: item.qcr_id, name: item.qcr_name, email: item.qcr_email || '' };
    } else {
        clearQcrSelection();
    }
}

// =============================================================================
// CONTRACTOR UPDATE REVIEW FUNCTIONS
// =============================================================================

/**
 * Render the contractor update review panel in the drawer
 */
function renderContractorUpdatePanel(item) {
    const container = document.getElementById('contractor-update-panel');
    if (!container) return;
    
    // Only show for admins and if there's a pending update
    if (!state.user || state.user.role !== 'admin' || !item.has_pending_update) {
        container.innerHTML = '';
        hide(container);
        return;
    }
    
    const isContentChange = item.update_type === 'content_change';
    const wasReopened = item.reopened_from_closed === 1;
    
    // Build change details
    let changesHtml = '';
    if (item.previous_due_date && item.due_date !== item.previous_due_date) {
        changesHtml += `
            <div class="update-change-item">
                <span class="label">Due Date:</span>
                <span class="old-value">${formatDate(item.previous_due_date)}</span>
                <span class="arrow">→</span>
                <span class="new-value">${formatDate(item.due_date)}</span>
            </div>
        `;
    }
    if (item.previous_title && item.title !== item.previous_title) {
        changesHtml += `
            <div class="update-change-item">
                <span class="label">Title:</span>
                <span class="old-value">${escapeHtml(item.previous_title?.substring(0, 50))}...</span>
                <span class="arrow">→</span>
                <span class="new-value">${escapeHtml(item.title?.substring(0, 50))}...</span>
            </div>
        `;
    }
    if (item.previous_priority && item.priority !== item.previous_priority) {
        changesHtml += `
            <div class="update-change-item">
                <span class="label">Priority:</span>
                <span class="old-value">${item.previous_priority}</span>
                <span class="arrow">→</span>
                <span class="new-value">${item.priority}</span>
            </div>
        `;
    }
    
    const panelClass = isContentChange ? 'content-change' : '';
    const headerIcon = isContentChange ? '⚠️' : '📅';
    const headerText = wasReopened 
        ? 'CLOSED ITEM UPDATED BY CONTRACTOR' 
        : (isContentChange ? 'CONTENT CHANGE FROM CONTRACTOR' : 'DUE DATE UPDATE FROM CONTRACTOR');
    
    container.innerHTML = `
        <div class="update-review-panel ${panelClass}">
            <div class="update-review-header">
                <span class="icon">${headerIcon}</span>
                <span>${headerText}</span>
                ${wasReopened ? '<span style="background:#dc2626;color:white;padding:2px 6px;border-radius:4px;font-size:0.75rem;margin-left:auto;">REOPENED</span>' : ''}
            </div>
            
            <div class="update-changes">
                ${changesHtml || '<div class="update-change-item"><span class="label">Update detected:</span> <span class="new-value">${formatDateTime(item.update_detected_at)}</span></div>'}
            </div>
            
            ${item.status_before_update ? `<p style="font-size:0.75rem;color:var(--text-secondary);margin-bottom:0.5rem;">Status before update: <strong>${item.status_before_update}</strong></p>` : ''}
            
            <div class="update-admin-note">
                <label style="font-size:0.75rem;font-weight:500;color:var(--text-secondary);">Admin Note (will be included in notification):</label>
                <textarea id="update-admin-note" placeholder="Describe the changes or provide context for reviewers..."></textarea>
            </div>
            
            <div class="update-actions">
                ${!isContentChange ? `
                    <button class="btn btn-accept-due-date" onclick="reviewContractorUpdate(${item.id}, 'accept_due_date')">
                        📅 Accept Due Date Change
                    </button>
                ` : ''}
                <button class="btn btn-restart-workflow" onclick="reviewContractorUpdate(${item.id}, 'restart_workflow')">
                    🔄 Restart Workflow to Reviewer
                </button>
                <button class="btn btn-dismiss-update" onclick="reviewContractorUpdate(${item.id}, 'dismiss')">
                    ✕ Dismiss
                </button>
            </div>
        </div>
    `;
    
    show(container);
}

/**
 * Handle admin review of contractor update
 */
async function reviewContractorUpdate(itemId, action) {
    const adminNote = document.getElementById('update-admin-note')?.value || '';
    
    let confirmMsg = '';
    switch (action) {
        case 'accept_due_date':
            confirmMsg = 'Accept the due date change? The appropriate party (reviewer or QCR) will be notified of the new due date.';
            break;
        case 'restart_workflow':
            confirmMsg = 'Restart the workflow? This will clear all responses and send new review requests to the reviewer(s).';
            break;
        case 'dismiss':
            confirmMsg = 'Dismiss this update without taking action?';
            break;
    }
    
    if (!confirm(confirmMsg)) return;
    
    try {
        const result = await api(`/item/${itemId}/review-update`, {
            method: 'POST',
            body: JSON.stringify({
                action: action,
                admin_note: adminNote,
                apply_new_values: true
            })
        });
        
        if (result.success) {
            let message = 'Update processed successfully.';
            if (result.emails_sent && result.emails_sent.length > 0) {
                const successEmails = result.emails_sent.filter(e => e.result?.success);
                message += ` Sent ${successEmails.length} notification(s).`;
            }
            showToast(message, 'success');
            
            // Refresh the item
            const item = await api(`/item/${itemId}`);
            populateDetailDrawer(item);
            
            // Refresh items list
            loadItems();
        } else {
            showToast(result.error || 'Failed to process update', 'error');
        }
    } catch (err) {
        console.error('Failed to review update:', err);
        showToast('Failed to process update: ' + err.message, 'error');
    }
}

/**
 * Load pending updates count for stats
 */
async function loadPendingUpdatesCount() {
    try {
        const updates = await api('/pending-updates');
        const count = updates.length;
        
        // Update the pending updates badge in sidebar
        const badge = document.getElementById('count-pending-updates');
        if (badge) {
            badge.textContent = count > 0 ? count : '';
            badge.style.display = count > 0 ? 'inline-flex' : 'none';
        }
        
        return count;
    } catch (err) {
        console.error('Failed to load pending updates:', err);
        return 0;
    }
}

// Global function aliases for onclick handlers in HTML
window.openItem = openDetailDrawer;
window.removeReviewerChip = removeReviewerChip;
window.addReviewerChip = addReviewerChip;
window.selectQcr = selectQcr;
window.reviewContractorUpdate = reviewContractorUpdate;