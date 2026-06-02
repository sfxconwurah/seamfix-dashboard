"""
Shared theme system for all Seamfix dashboards.
Provides CSS variables for light/dark mode and a toggle button.
Light mode is the default. User preference is saved to localStorage.
"""

# CSS variable definitions — light mode default, dark mode override
THEME_CSS = """
:root {
  --bg-body: #f1f5f9;
  --bg-card: #ffffff;
  --bg-card-hover: rgba(0,212,170,0.06);
  --bg-nav: #ffffff;
  --bg-input: #ffffff;
  --bg-table-header: rgba(0,150,120,0.08);
  --bg-table-hover: rgba(0,150,120,0.04);
  --bg-table-stripe: rgba(0,0,0,0.02);
  --bg-gauge: rgba(0,0,0,0.08);
  --text-primary: #0f172a;
  --text-secondary: #334155;
  --text-tertiary: #64748b;
  --text-heading: #1e293b;
  --text-on-accent: #ffffff;
  --border-main: #e2e8f0;
  --border-light: #f1f5f9;
  --border-accent: rgba(0,180,150,0.25);
  --accent: #009E7E;
  --accent-bright: #00D4AA;
  --accent-secondary: #3BA89E;
  --accent-bg: rgba(0,212,170,0.08);
  --warning: #b45309;
  --warning-bright: #f59e0b;
  --warning-bg: rgba(245,158,11,0.10);
  --danger: #dc2626;
  --danger-bright: #FF6B6B;
  --danger-bg: rgba(255,107,107,0.08);
  --success-bg: rgba(0,212,170,0.06);
  --chart-grid: rgba(0,0,0,0.06);
  --chart-text: #475569;
  --shadow-sm: 0 1px 3px rgba(0,0,0,0.06);
  --shadow-hover: 0 4px 12px rgba(0,0,0,0.08);
  --print-bg: #fff;
}

[data-theme="dark"] {
  --bg-body: #0f172a;
  --bg-card: #1e293b;
  --bg-card-hover: rgba(0,212,170,0.08);
  --bg-nav: #0a0f1e;
  --bg-input: rgba(15,23,42,0.8);
  --bg-table-header: rgba(0,212,170,0.08);
  --bg-table-hover: rgba(0,212,170,0.04);
  --bg-table-stripe: rgba(255,255,255,0.02);
  --bg-gauge: rgba(0,0,0,0.3);
  --text-primary: #e2e8f0;
  --text-secondary: #94a3b8;
  --text-tertiary: #64748b;
  --text-heading: #cbd5e1;
  --text-on-accent: #ffffff;
  --border-main: #334155;
  --border-light: rgba(0,212,170,0.06);
  --border-accent: rgba(0,212,170,0.15);
  --accent: #00D4AA;
  --accent-bright: #00D4AA;
  --accent-secondary: #4ECDC4;
  --accent-bg: rgba(0,212,170,0.08);
  --warning: #FFE66D;
  --warning-bright: #FFE66D;
  --warning-bg: rgba(255,230,109,0.10);
  --danger: #FF6B6B;
  --danger-bright: #FF6B6B;
  --danger-bg: rgba(255,107,107,0.05);
  --success-bg: rgba(0,212,170,0.06);
  --chart-grid: rgba(0,212,170,0.08);
  --chart-text: #64748b;
  --shadow-sm: 0 1px 3px rgba(0,0,0,0.2);
  --shadow-hover: 0 8px 30px rgba(0,212,170,0.1);
  --print-bg: #fff;
}
"""

# Toggle button HTML — floating button, NOT inside top-nav (Streamlit hides top-nav)
THEME_TOGGLE_HTML = """<button onclick="toggleTheme()" id="themeToggle" style="position:fixed;top:12px;right:16px;z-index:9999;background:var(--bg-card);border:1px solid var(--border-main);border-radius:8px;padding:6px 14px;cursor:pointer;color:var(--text-secondary);font-size:12px;font-family:inherit;display:flex;align-items:center;gap:5px;height:34px;white-space:nowrap;transition:all 0.2s;box-shadow:0 2px 8px rgba(0,0,0,0.1)" onmouseover="this.style.borderColor='var(--accent)';this.style.color='var(--accent)'" onmouseout="this.style.borderColor='var(--border-main)';this.style.color='var(--text-secondary)'" title="Toggle light/dark mode">
<span id="themeIcon">&#9790;</span> <span id="themeLabel">Dark</span>
</button>"""

# JavaScript for theme toggle + localStorage persistence
THEME_JS = """
function toggleTheme() {
  var current = document.documentElement.getAttribute('data-theme');
  var next = current === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', next);
  localStorage.setItem('seamfix-theme', next);
  updateToggleUI(next);
  updateCharts(next);
  try { window.parent.postMessage({seamfixTheme: next}, '*'); } catch(e) {}
}

function updateToggleUI(theme) {
  var icon = document.getElementById('themeIcon');
  var label = document.getElementById('themeLabel');
  if (icon && label) {
    icon.innerHTML = theme === 'dark' ? '&#9788;' : '&#9790;';
    label.textContent = theme === 'dark' ? 'Light' : 'Dark';
  }
}

function updateCharts(theme) {
  if (typeof Chart === 'undefined') return;
  var gridColor = theme === 'dark' ? 'rgba(0,212,170,0.08)' : 'rgba(0,0,0,0.06)';
  var tickColor = theme === 'dark' ? '#64748b' : '#64748b';
  var legendColor = theme === 'dark' ? '#94a3b8' : '#475569';
  Chart.helpers.each(Chart.instances, function(instance) {
    if (instance.config && instance.config.options) {
      var opts = instance.config.options;
      if (opts.scales) {
        ['x','y'].forEach(function(axis) {
          if (opts.scales[axis]) {
            if (opts.scales[axis].grid) opts.scales[axis].grid.color = gridColor;
            if (opts.scales[axis].ticks) opts.scales[axis].ticks.color = tickColor;
          }
        });
      }
      if (opts.plugins && opts.plugins.legend && opts.plugins.legend.labels) {
        opts.plugins.legend.labels.color = legendColor;
      }
      instance.update('none');
    }
  });
}

(function() {
  var saved = localStorage.getItem('seamfix-theme') || 'light';
  document.documentElement.setAttribute('data-theme', saved);
  updateToggleUI(saved);
})();
"""


def get_base_css():
    """Return the complete theme CSS variable block."""
    return THEME_CSS


def get_toggle_html():
    """Return the HTML for the theme toggle button (goes in top-nav)."""
    return THEME_TOGGLE_HTML


def get_theme_js():
    """Return the JavaScript for theme switching."""
    return THEME_JS
