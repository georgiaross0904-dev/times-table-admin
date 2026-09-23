(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.TimesTableQuiz = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const MODES = Object.freeze({ MIXED: 'mixed', MULTIPLY: 'multiply', DIVIDE: 'divide' });

  function integerBetween(min, max, rng) {
    return Math.floor(rng() * (max - min + 1)) + min;
  }

  function multiplicationQuestion(rng = Math.random) {
    const left = integerBetween(0, 12, rng);
    const right = integerBetween(0, 12, rng);
    return {
      type: 'multiply',
      left,
      right,
      answer: left * right,
      text: `${left} × ${right} = ?`,
      answerText: `Answer: ${left * right}`,
      key: `m:${left}:${right}`
    };
  }

  function divisionQuestion(rng = Math.random) {
    const divisor = integerBetween(1, 12, rng); // Deliberately never zero.
    const quotient = integerBetween(0, 12, rng);
    const dividend = divisor * quotient;
    return {
      type: 'divide',
      left: dividend,
      right: divisor,
      answer: quotient,
      text: `${dividend} ÷ ${divisor} = ?`,
      answerText: `Answer: ${quotient}`,
      key: `d:${dividend}:${divisor}`
    };
  }

  function generateQuestion(mode = MODES.MIXED, rng = Math.random, previousKey = null) {
    let question;
    for (let attempt = 0; attempt < 12; attempt += 1) {
      if (mode === MODES.MULTIPLY) question = multiplicationQuestion(rng);
      else if (mode === MODES.DIVIDE) question = divisionQuestion(rng);
      else question = rng() < 0.5 ? multiplicationQuestion(rng) : divisionQuestion(rng);
      if (question.key !== previousKey) return question;
    }
    return question;
  }

  function durationToMs(minutes) {
    const value = Number(minutes);
    if (!Number.isFinite(value) || value <= 0) return null;
    return value * 60 * 1000;
  }

  function formatRemaining(ms) {
    if (ms === null) return 'Unlimited';
    const safe = Math.max(0, Math.ceil(ms / 1000));
    const minutes = Math.floor(safe / 60);
    const seconds = safe % 60;
    return `${minutes}:${String(seconds).padStart(2, '0')}`;
  }

  return { MODES, multiplicationQuestion, divisionQuestion, generateQuestion, durationToMs, formatRemaining };
});
