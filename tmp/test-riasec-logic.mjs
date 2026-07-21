import fs from 'node:fs/promises';
import vm from 'node:vm';

const source = await fs.readFile('app.js', 'utf8');
const start = source.indexOf('const riasecOrder=');
const end = source.indexOf('function findCareerMatches');
const context = { codes: { R:{}, I:{}, A:{}, S:{}, E:{}, C:{} } };
vm.runInNewContext(source.slice(start, end), context);
const { calculateRiasecResult } = context;

function questionsFor(scores) {
  return Object.entries(scores).flatMap(([code, count]) => Array.from({ length: count }, () => ({ code, type: 'อาชีพ' })));
}
const primary = calculateRiasecResult(questionsFor({ R: 22, I: 15, A: 8, S: 30, E: 25, C: 12 }), Array(112).fill(1));
if (primary.summaryCode !== 'SER') throw new Error(`Expected SER, got ${primary.summaryCode}`);

const tieQuestions = [
  ...Array.from({ length: 20 }, () => ({ code: 'S', type: 'อาชีพ' })),
  ...Array.from({ length: 10 }, () => ({ code: 'E', type: 'อาชีพ' })),
  ...Array.from({ length: 10 }, () => ({ code: 'C', type: 'อาชีพ' })),
  { code: 'R', type: 'กิจกรรม' },
  { code: 'I', type: 'กิจกรรม' },
  { code: 'A', type: 'กิจกรรม' },
];
const tieAnswers = Array(tieQuestions.length).fill(1);
const tie = calculateRiasecResult(tieQuestions, tieAnswers);
if (tie.summaryCode !== 'SEC' || !tie.alternativeCodes.includes('SCE')) throw new Error(`Unexpected tie result: ${JSON.stringify(tie)}`);

console.log(JSON.stringify({ primary: primary.summaryCode, tie: tie.summaryCode, alternatives: tie.alternativeCodes }));
