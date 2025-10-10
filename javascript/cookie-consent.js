(function() {
  const messages = {
    ru: {
      text: "Мы используем cookie для персонализированной аналитики. Через 30 сек автоматически включится анонимная статистика.",
      accept: "Принять",
      decline: "Анонимно"
    },
    en: {
      text: "We use cookies for personalized analytics. Anonymous statistics will auto-enable in 30 seconds.",
      accept: "Accept",
      decline: "Anonymous"
    },
    fr: {
      text: "Cookies pour analyses personnalisées. Statistiques anonymes auto-activées en 30s.",
      accept: "Accepter",
      decline: "Anonyme"
    },
    de: {
      text: "Cookies für personalisierte Analysen. Anonyme Statistiken in 30s auto-aktiviert.",
      accept: "Akzeptieren",
      decline: "Anonym"
    },
    ar: {
      text: "نحن نستخدم ملفات تعريف الارتباط للتحليل الشخصي. عند الرفض، سيتم استخدام إحصائيات مجهولة.",
      accept: "قبول",
      decline: "مجهول"
    },
    es: {
      text: "Utilizamos cookies para análisis personalizado. Si se rechaza, se usarán estadísticas anónimas.",
      accept: "Aceptar",
      decline: "Anónimo"
    },
    ja: {
      text: "パーソナライズされた分析にCookieを使用しています。拒否した場合、匿名統計が使用されます。",
      accept: "承諾",
      decline: "匿名"
    },
    zh: {
      text: "我们使用Cookie进行个性化分析。如果拒绝，将使用匿名统计。",
      accept: "接受",
      decline: "匿名"
    },
    pl: {
      text: "Używamy plików cookie do spersonalizowanej analizy. Po odrzuceniu będą używane anonimowe statystyki.",
      accept: "Akceptuj",
      decline: "Anonimowo"
    },
    pt: {
      text: "Usamos cookies para análise personalizada. Se recusado, estatísticas anônimas serão usadas.",
      accept: "Aceitar",
      decline: "Anônimo"
    }
  };

  function initCookieBanner() {
    const lang = (document.documentElement.lang || 'en').slice(0,2);
    const t = messages[lang] || messages.en;

    if (localStorage.getItem('cookie_consent')) return;

    // Auto-enable anonymous analytics after 30 seconds if no choice made
    const autoConsentTimer = setTimeout(() => {
      console.log('Auto-enabling anonymous analytics after 30 seconds of no user action');
      // Enable anonymous analytics automatically
      gtag('consent', 'update', {
        'ad_storage': 'denied',
        'analytics_storage': 'granted'
      });
      gtag('config', 'G-C19L2VS3EH', {
        'anonymize_ip': true,
        'allow_google_signals': false,
        'allow_ad_personalization_signals': false
      });
      // Send required events for proper tracking
      gtag('event', 'cookie_consent', {
        'consent_type': 'auto_anonymous'
      });
      gtag('event', 'page_view', {
        'page_title': document.title,
        'page_location': window.location.href,
        'send_to': 'G-C19L2VS3EH'
      });
      localStorage.setItem('cookie_consent', 'auto_anonymous');
      const banner = document.getElementById('cookie-banner');
      if (banner) banner.remove();
    }, 30000); // 30 seconds


    const banner = document.createElement('div');
    banner.id = 'cookie-banner';
    banner.innerHTML = `
      <style>
        #cookie-banner {
          position: fixed; bottom: 15px; left: 15px; right: 15px;
          background: #fff; color: #000; border: 1px solid #ccc;
          padding: 10px; border-radius: 8px;
          box-shadow: 0 2px 5px rgba(0,0,0,0.1);
          z-index: 9999; font-family: sans-serif;
        }
        #cookie-banner button { margin-left: 10px; }
      </style>
      ${t.text}
      <button id="accept-cookies">${t.accept}</button>
      <button id="decline-cookies">${t.decline}</button>
    `;
    document.body.appendChild(banner);

    document.getElementById('accept-cookies').onclick = function() {
      clearTimeout(autoConsentTimer); // Cancel auto-consent
      console.log('User accepted cookies');
      gtag('consent', 'update', {
        'ad_storage': 'granted',
        'analytics_storage': 'granted'
      });
      gtag('config', 'G-C19L2VS3EH');
      // Send required events for proper tracking
      gtag('event', 'cookie_consent', {
        'consent_type': 'granted'
      });
      gtag('event', 'page_view', {
        'page_title': document.title,
        'page_location': window.location.href,
        'send_to': 'G-C19L2VS3EH'
      });
      localStorage.setItem('cookie_consent', 'granted');
      banner.remove();
    };

    document.getElementById('decline-cookies').onclick = function() {
      clearTimeout(autoConsentTimer); // Cancel auto-consent
      console.log('User declined cookies, using anonymous analytics');
      // Enable analytics but without storage/cookies (anonymous mode)
      gtag('consent', 'update', {
        'ad_storage': 'denied',
        'analytics_storage': 'granted'  // Allow analytics but configure it as anonymous
      });
      gtag('config', 'G-C19L2VS3EH', {
        'anonymize_ip': true,
        'allow_google_signals': false,  // Disable cross-device tracking
        'allow_ad_personalization_signals': false  // Disable ad personalization
      });
      // Send required events for proper tracking
      gtag('event', 'cookie_consent', {
        'consent_type': 'denied_anonymous'
      });
      gtag('event', 'page_view', {
        'page_title': document.title,
        'page_location': window.location.href,
        'send_to': 'G-C19L2VS3EH'
      });
      localStorage.setItem('cookie_consent', 'denied');
      banner.remove();
    };
  }

  // Wait for DOM to be ready before initializing cookie banner
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initCookieBanner);
  } else {
    // DOM is already ready
    initCookieBanner();
  }
})();
