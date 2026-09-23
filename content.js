(function () {
  'use strict';

  var initialised = false;
  var state = null;
  var timerId = null;

  var core;
  var app;
  var setupScreen;
  var quizScreen;
  var startButton;
  var pauseButton;
  var resetButton;
  var questionSeconds;
  var answerSeconds;
  var durationMinutes;
  var questionText;
  var answerText;
  var statusText;
  var timeRemaining;
  var phaseCountdown;
  var runtimeStatus;

  function byId(id) { return document.getElementById(id); }

  function showRuntimeStatus(message, isError) {
    if (!runtimeStatus) return;
    runtimeStatus.textContent = message || '';
    runtimeStatus.className = isError ? 'runtime-status is-error' : 'runtime-status';
  }

  function selectedMode() {
    var selected = document.querySelector('input[name="mode"]:checked');
    return selected ? selected.value : 'mixed';
  }

  function nowMs() {
    return (window.performance && typeof window.performance.now === 'function')
      ? window.performance.now()
      : Date.now();
  }

  function newState() {
    var durationMs = core.durationToMs(durationMinutes.value);
    return {
      mode: selectedMode(),
      questionMs: Number(questionSeconds.value) * 1000,
      answerMs: Number(answerSeconds.value) * 1000,
      durationMs: durationMs,
      sessionRemainingMs: durationMs,
      phase: 'question',
      phaseRemainingMs: Number(questionSeconds.value) * 1000,
      currentQuestion: null,
      previousKey: null,
      paused: false,
      ended: false,
      lastTickAt: nowMs()
    };
  }

  function showNewQuestion() {
    state.currentQuestion = core.generateQuestion(state.mode, Math.random, state.previousKey);
    state.previousKey = state.currentQuestion.key;
    state.phase = 'question';
    state.phaseRemainingMs = state.questionMs;
    questionText.textContent = state.currentQuestion.text;
    answerText.textContent = '';
    render();
  }

  function revealAnswer() {
    state.phase = 'answer';
    state.phaseRemainingMs = state.answerMs;
    answerText.textContent = state.currentQuestion.answerText;
    render();
  }

  function stopTimer() {
    if (timerId !== null) {
      window.clearInterval(timerId);
      timerId = null;
    }
  }

  function endSession() {
    state.ended = true;
    state.paused = false;
    app.classList.remove('is-paused');
    app.classList.add('is-ended');
    statusText.textContent = 'Complete';
    timeRemaining.textContent = '0:00';
    questionText.textContent = 'Session complete';
    answerText.textContent = '';
    phaseCountdown.textContent = '';
    pauseButton.textContent = 'Restart';
    stopTimer();
  }

  function consumeElapsed(elapsed) {
    if (!state || state.paused || state.ended) return;

    if (state.sessionRemainingMs !== null) {
      state.sessionRemainingMs -= elapsed;
      if (state.sessionRemainingMs <= 0) {
        state.sessionRemainingMs = 0;
        endSession();
        return;
      }
    }

    var remainingElapsed = elapsed;
    while (remainingElapsed >= state.phaseRemainingMs && !state.ended) {
      remainingElapsed -= state.phaseRemainingMs;
      if (state.phase === 'question') revealAnswer();
      else showNewQuestion();
    }
    state.phaseRemainingMs -= remainingElapsed;
  }

  function tick() {
    if (!state) return;
    var now = nowMs();
    var elapsed = Math.max(0, now - state.lastTickAt);
    state.lastTickAt = now;
    consumeElapsed(elapsed);
    render();
  }

  function startTimer() {
    stopTimer();
    state.lastTickAt = nowMs();
    timerId = window.setInterval(tick, 100);
  }

  function render() {
    if (!state) return;
    statusText.textContent = state.ended ? 'Complete' : state.paused ? 'Paused' : (state.phase === 'answer' ? 'Answer' : 'Question');
    timeRemaining.textContent = core.formatRemaining(state.sessionRemainingMs);
    phaseCountdown.textContent = state.ended ? '' : (Math.max(0, state.phaseRemainingMs / 1000).toFixed(1) + 's');
    pauseButton.textContent = state.ended ? 'Restart' : state.paused ? 'Resume' : 'Pause';
  }

  function startQuiz() {
    try {
      state = newState();
      app.classList.remove('is-paused', 'is-ended');
      setupScreen.hidden = true;
      quizScreen.hidden = false;
      showNewQuestion();
      startTimer();
    } catch (error) {
      setupScreen.hidden = false;
      quizScreen.hidden = true;
      showRuntimeStatus('Could not start: ' + (error && error.message ? error.message : String(error)), true);
    }
  }

  function togglePause() {
    if (!state) return;
    if (state.ended) {
      startQuiz();
      return;
    }
    state.paused = !state.paused;
    app.classList.toggle('is-paused', state.paused);
    state.lastTickAt = nowMs();
    render();
  }

  function resetQuiz() {
    stopTimer();
    state = null;
    app.classList.remove('is-paused', 'is-ended');
    quizScreen.hidden = true;
    setupScreen.hidden = false;
    questionText.textContent = '8 × 7 = ?';
    answerText.textContent = '';
    showRuntimeStatus('Ready', false);
  }

  function wireControls() {
    if (initialised) return;
    initialised = true;

    core = window.TimesTableQuiz;
    app = byId('app');
    setupScreen = byId('setupScreen');
    quizScreen = byId('quizScreen');
    startButton = byId('startButton');
    pauseButton = byId('pauseButton');
    resetButton = byId('resetButton');
    questionSeconds = byId('questionSeconds');
    answerSeconds = byId('answerSeconds');
    durationMinutes = byId('durationMinutes');
    questionText = byId('questionText');
    answerText = byId('answerText');
    statusText = byId('statusText');
    timeRemaining = byId('timeRemaining');
    phaseCountdown = byId('phaseCountdown');
    runtimeStatus = byId('runtimeStatus');

    if (!core || !startButton) {
      showRuntimeStatus('Quiz files did not load correctly.', true);
      return;
    }

    startButton.addEventListener('click', startQuiz);
    pauseButton.addEventListener('click', togglePause);
    resetButton.addEventListener('click', resetQuiz);
    startButton.disabled = false;
    startButton.textContent = 'Start';
    showRuntimeStatus('Ready', false);
  }

  function startWhenOfficeReady() {
    if (window.Office && typeof window.Office.onReady === 'function') {
      window.Office.onReady(function () {
        wireControls();
      }).catch(function (error) {
        wireControls();
        showRuntimeStatus('PowerPoint runtime warning: ' + (error && error.message ? error.message : String(error)), true);
      });
    } else {
      /* This fallback also makes the page testable in an ordinary browser. */
      wireControls();
      showRuntimeStatus('Ready (browser fallback)', false);
    }
  }

  window.addEventListener('error', function (event) {
    if (runtimeStatus) showRuntimeStatus('Script error: ' + event.message, true);
  });

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', startWhenOfficeReady);
  } else {
    startWhenOfficeReady();
  }
})();
