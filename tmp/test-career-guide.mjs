import fs from 'node:fs/promises';
import vm from 'node:vm';

const source = await fs.readFile('career-guide-data.js', 'utf8');
const context = {};
vm.runInNewContext(`${source}\nthis.search = findCareerGuides; this.guides = careerGuides;`, context);

if (context.guides.length < 20) throw new Error('Expected at least 20 career guides');
if (context.search('หมอ')[0]?.name !== 'แพทย์') throw new Error('Alias search for หมอ failed');
if (context.search('โปรแกรมเมอร์')[0]?.name !== 'นักพัฒนาซอฟต์แวร์') throw new Error('Alias search for โปรแกรมเมอร์ failed');
if (context.search('UX Designer')[0]?.name !== 'นักออกแบบ UX/UI') throw new Error('Case-insensitive English search failed');
if (context.search('อาชีพที่ไม่มีในระบบ').length !== 0) throw new Error('Unknown career should not match');

console.log(JSON.stringify({ guides: context.guides.length, aliases: 'ok', unknown: 'ok' }));
