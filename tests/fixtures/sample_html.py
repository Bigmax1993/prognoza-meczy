SAMPLE_MATCHES_HTML = """<!DOCTYPE html>
<html>
<head><title>FootyStats Test</title></head>
<body>
  <h2 class="postTitle dark-gray fl">- 11/08</h2>

  <div class="league">
    <div class="league-header">
      <span class="league-country">Европа -</span>
      <span class="league-title">Европа - Лига чемпионов УЕФА</span>
    </div>
    <div class="league-matches">
      <a class="match row cf z1" href="/ru/europe/team-a-vs-team-b-h2h-stats">
        <div class="date convert-months time" data-time="1786467600">
          <div class="used-to-be-a">
            <span class="timezone-convert-match-regular" data-time="1786467600">18:00</span>
            <span class="match-time-soon" data-match-status="incomplete" data-match-time="1786467600"></span>
          </div>
        </div>
        <div class="match-info row cf fl rfnone">
          <div class="team home fl">
            <div class="fr">
              <span class="hover-modal-parent hover-modal-ajax-team" data-team-id="1">Жальгирис</span>
              <div class="form-box okay1">1.60</div>
            </div>
          </div>
          <div class="h2h-link pr fl"><span class="ft-indicator blue">Стат-ка</span></div>
          <div class="team away fl">
            <div class="form-box excellent">2.33</div>
            <span class="hover-modal-parent hover-modal-ajax-team" data-team-id="2">Динамо Загреб</span>
          </div>
        </div>
        <div class="match-stats fl">
          <div class="stat odds dark-gray bbox">
            <span class="col-lg-4 col-sm-4 ac hover-modal-parent">6.35
              <span class="hover-modal-content"><span class="semi-bold">Home Win</span></span>
            </span>
            <span class="col-lg-4 col-sm-4 ac hover-modal-parent">4.62
              <span class="hover-modal-content"><span class="semi-bold">Draw</span></span>
            </span>
            <span class="col-lg-4 col-sm-4 ac hover-modal-parent">1.39
              <span class="hover-modal-content"><span class="semi-bold">Away Win</span></span>
            </span>
          </div>
        </div>
      </a>
      <a class="match row cf z2" href="/ru/europe/nec-vs-olympiakos-h2h-stats">
        <div class="date convert-months time">
          <div class="used-to-be-a">
            <span class="timezone-convert-match-regular">18:30</span>
            <span class="match-time-soon" data-match-status="incomplete"></span>
          </div>
        </div>
        <div class="match-info row cf fl rfnone">
          <div class="team home fl">
            <div class="fr">
              <span class="hover-modal-parent hover-modal-ajax-team">НЕК</span>
              <div class="form-box">1.00</div>
            </div>
          </div>
          <div class="team away fl">
            <div class="form-box">1.00</div>
            <span class="hover-modal-parent hover-modal-ajax-team">Олимпиакос</span>
          </div>
        </div>
        <div class="match-stats fl">
          <div class="stat odds dark-gray bbox">
            <span class="col-lg-4 col-sm-4 ac hover-modal-parent">2.75</span>
            <span class="col-lg-4 col-sm-4 ac hover-modal-parent">3.40</span>
            <span class="col-lg-4 col-sm-4 ac hover-modal-parent">2.45</span>
          </div>
        </div>
      </a>
    </div>
  </div>

  <div class="league">
    <div class="league-header">
      <span class="league-country">Англия -</span>
      <span class="league-title">Англия - Кубок Англии</span>
    </div>
    <div class="league-matches">
      <a class="match row cf z3" href="/ru/england/hallen-vs-roman-h2h-stats">
        <div class="date">
          <span class="timezone-convert-match-regular">19:30</span>
          <span class="match-time-soon" data-match-status="incomplete"></span>
        </div>
        <div class="match-info">
          <div class="team home fl">
            <div class="fr">
              <span class="hover-modal-parent hover-modal-ajax-team">Холлен</span>
              <div class="form-box">1.00</div>
            </div>
          </div>
          <div class="team away fl">
            <div class="form-box">1.00</div>
            <span class="hover-modal-parent hover-modal-ajax-team">Роман Гласс</span>
          </div>
        </div>
        <div class="match-stats fl">
          <div class="stat odds dark-gray bbox"></div>
        </div>
      </a>
    </div>
  </div>
</body>
</html>
"""

CLOUDFLARE_HTML = """<!DOCTYPE html>
<html><head><title>Just a moment...</title></head>
<body>
  <h1>footystats.org</h1>
  <p>Checking your browser</p>
  <script src="https://challenges.cloudflare.com/turnstile/v0/api.js"></script>
  <div class="cf-turnstile"></div>
</body></html>
"""

CLOUDFLARE_RU_HTML = """<!DOCTYPE html>
<html><head><title>Один момент…</title></head>
<body>
  <h2>Выполнение проверки безопасности</h2>
  <div id="challenge-platform"></div>
</body></html>
"""
