(function() {
  const messages = {
    ru: {
      text: "Мы используем cookie для персонализированной аналитики. При отклонении будет использоваться анонимная статистика.",
      accept: "Принять",
      decline: "Анонимно"
    },
    en: {
      text: "We use cookies for personalized analytics. If declined, anonymous statistics will be used.",
      accept: "Accept",
      decline: "Anonymous"
    },
    fr: {
      text: "Nous utilisons des cookies pour l'analyse personnalisée. Si refusé, des statistiques anonymes seront utilisées.",
      accept: "Accepter",
      decline: "Anonyme"
    },
    de: {
      text: "Wir verwenden Cookies für personalisierte Analysen. Bei Ablehnung werden anonyme Statistiken verwendet.",
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
      gtag('consent', 'update', {
        'ad_storage': 'granted',
        'analytics_storage': 'granted'
      });
      localStorage.setItem('cookie_consent', 'granted');
      banner.remove();
    };

    document.getElementById('decline-cookies').onclick = function() {
      gtag('consent', 'update', {
        'ad_storage': 'denied',
        'analytics_storage': 'denied'
      });
      // Enable anonymous analytics without cookies
      gtag('config', 'G-C19L2VS3EH', {
        'anonymize_ip': true,
        'storage': 'none',
        'client_storage': 'none'
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
