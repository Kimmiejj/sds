import fs from 'node:fs/promises';
import vm from 'node:vm';

const officialSource = await fs.readFile('thai-occupations-data.js', 'utf8');
const source = await fs.readFile('career-guide-data.js', 'utf8');
const context = {};
vm.runInNewContext(`${officialSource}\n${source}\nthis.search = findCareerGuides; this.guides = careerGuides; this.official = thaiOccupations; this.searchOfficial = findThaiOccupations;`, context);

if (context.guides.length < 20) throw new Error('Expected at least 20 career guides');
if (context.official.length < 400) throw new Error('Expected at least 400 official Thai occupations');
if (context.search('หมอ')[0]?.name !== 'แพทย์') throw new Error('Alias search for หมอ failed');
if (context.search('โปรแกรมเมอร์')[0]?.name !== 'นักพัฒนาซอฟต์แวร์') throw new Error('Alias search for โปรแกรมเมอร์ failed');
if (context.search('UX Designer')[0]?.name !== 'นักออกแบบ UX/UI') throw new Error('Case-insensitive English search failed');
if (context.search('อาชีพที่ไม่มีในระบบ').length !== 0) throw new Error('Unknown career should not match');
if (context.searchOfficial('ช่างเทคนิคด้านวิศวกรรมไฟฟ้า')[0]?.group !== 3) throw new Error('Official occupation group search failed');

console.log(JSON.stringify({ guides: context.guides.length, official: context.official.length, aliases: 'ok', unknown: 'ok' }));
