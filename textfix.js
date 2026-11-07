/* textfix.js v1.0.0 — EcoOils HSSE Recognition: offline text tidy-up. No AI, no internet, no key.
   Fixes spelling of common site words, SMS / informal words (EN, Bahasa Indonesia, Malay), spacing, punctuation,
   capital letters and a few weak phrases. It never adds facts and never translates.
   Use:  TextFix.fix('text')  ->  { text, count, lang }      TextFix.fixAll({a:'..', b:'..'})  ->  { text:{a,b}, count, lang }
   Add words: put them in the lists below (format  ['wrong', 'right']). Words are matched whole, any capitals. */
(function (root) {
  'use strict';
  var W = "\\p{L}\\p{N}'’";                       // characters that belong to a word
  function esc(s) { return s.replace(/[.*+?^${}()|[\]\\\/]/g, '\\$&'); }
  function compile(pairs) {
    var map = {}; pairs.forEach(function (p) { map[p[0].toLowerCase()] = p[1]; });
    var keys = Object.keys(map).sort(function (a, b) { return b.length - a.length; });
    return { map: map, re: new RegExp('(^|[^' + W + '])(' + keys.map(esc).join('|') + ')(?=$|[^' + W + '])', 'giu') };
  }

  /* ---- site words, wrong spelling -> right spelling (all languages) ---- */
  var SITE = compile([
    ['scafold', 'scaffold'], ['scafolds', 'scaffolds'], ['scafolding', 'scaffolding'], ['scaffolfing', 'scaffolding'], ['scaffoldin', 'scaffolding'], ['scafoldin', 'scaffolding'],
    ['housekping', 'housekeeping'], ['housekeping', 'housekeeping'], ['housekeeing', 'housekeeping'], ['housekepping', 'housekeeping'], ['house keeping', 'housekeeping'],
    ['harnes', 'harness'], ['harnass', 'harness'], ['extingusher', 'extinguisher'], ['extinguiser', 'extinguisher'], ['extinguiher', 'extinguisher'],
    ['hazzard', 'hazard'], ['hazzards', 'hazards'], ['hazrd', 'hazard'], ['helment', 'helmet'], ['helmat', 'helmet'], ['hellmet', 'helmet'],
    ['electrik', 'electric'], ['electical', 'electrical'], ['electircal', 'electrical'], ['barikade', 'barricade'], ['barracade', 'barricade'], ['baricade', 'barricade'],
    ['equiptment', 'equipment'], ['equipement', 'equipment'], ['equipmet', 'equipment'], ['cabel', 'cable'], ['cabble', 'cable'], ['cabels', 'cables'],
    ['terminaton', 'termination'], ['excavtion', 'excavation'], ['excavaton', 'excavation'], ['confind', 'confined'], ['crain', 'crane'], ['forklif', 'forklift'], ['forkift', 'forklift'],
    ['walk way', 'walkway'], ['work shop', 'workshop'], ['hand rail', 'handrail'], ['safty', 'safety'], ['saftey', 'safety'], ['safetly', 'safety'], ['unsafty', 'unsafe'],
    ['toolbox', 'toolbox'], ['tool box', 'toolbox'], ['spil', 'spill'], ['spilage', 'spillage'], ['lekage', 'leakage'], ['leackage', 'leakage'],
    ['berserakkan', 'berserakan'], ['resiko', 'risiko'], ['aktifitas', 'aktivitas'], ['praktek', 'praktik'], ['ijin', 'izin'], ['sistim', 'sistem'], ['kwalitas', 'kualitas'],
    ['merubah', 'mengubah'], ['dirubah', 'diubah'], ['merapihkan', 'merapikan'], ['dirapihkan', 'dirapikan'], ['rapih', 'rapi'], ['dikarenakan', 'karena'], ['apotik', 'apotek']
  ]);
  /* ---- acronyms and names written in capitals (all languages) ---- */
  var ACR = compile([
    ['ppe', 'PPE'], ['ptw', 'PTW'], ['jha', 'JHA'], ['jms', 'JMS'], ['loto', 'LOTO'], ['scba', 'SCBA'], ['mcc', 'MCC'], ['dtp', 'DTP'], ['etp', 'ETP'], ['wwtp', 'WWTP'], ['wtp', 'WTP'],
    ['ptu', 'PTU'], ['hsse', 'HSSE'], ['hse', 'HSE'], ['uauc', 'UAUC'], ['pln', 'PLN'], ['vam', 'VAM'], ['apd', 'APD'], ['k3', 'K3'],
    ['ecooils', 'EcoOils'], ['gsb', 'GSB'], ['kks', 'KKS'], ['pbs', 'PBS'], ['wen', 'WEN']
  ]);
  /* ---- English: SMS and common spelling ---- */
  var EN = compile([
    ['teh', 'the'], ['adn', 'and'], ['taht', 'that'], ['wich', 'which'], ['becuase', 'because'], ['becasue', 'because'], ['beacuse', 'because'], ['bcos', 'because'], ['bcz', 'because'], ['cuz', 'because'],
    ['recieve', 'receive'], ['recieved', 'received'], ['occured', 'occurred'], ['occurence', 'occurrence'], ['seperate', 'separate'], ['seperated', 'separated'], ['untill', 'until'],
    ['definately', 'definitely'], ['accomodate', 'accommodate'], ['begining', 'beginning'], ['beleive', 'believe'], ['comming', 'coming'], ['immediatly', 'immediately'],
    ['neccessary', 'necessary'], ['necesary', 'necessary'], ['noticable', 'noticeable'], ['posible', 'possible'], ['recomend', 'recommend'], ['responsability', 'responsibility'],
    ['succesful', 'successful'], ['tommorow', 'tomorrow'], ['wether', 'whether'], ['writting', 'writing'], ['alot', 'a lot'],
    ['dont', "don't"], ['cant', "can't"], ['wont', "won't"], ['didnt', "didn't"], ['doesnt', "doesn't"], ['wasnt', "wasn't"], ['isnt', "isn't"], ['arent', "aren't"], ['werent', "weren't"],
    ['couldnt', "couldn't"], ['shouldnt', "shouldn't"], ['wouldnt', "wouldn't"], ['hasnt', "hasn't"], ['havent', "haven't"], ['im', "I'm"], ['ive', "I've"],
    ['thx', 'thanks'], ['pls', 'please'], ['plz', 'please'], ['u', 'you'], ['ur', 'your'], ['b4', 'before'], ['w/o', 'without'], ['w/', 'with'], ['tho', 'though'], ['thru', 'through']
  ]);
  /* ---- English: weak phrase -> stronger phrase (meaning stays the same) ---- */
  var EN_PHR = compile([
    ['saw that', 'observed that'], ['saw a', 'observed a'], ['saw an', 'observed an'], ['saw the', 'observed the'], ['noticed that', 'observed that'], ['noticed a', 'observed a'], ['noticed the', 'observed the'],
    ['found out that', 'found that'], ['took out', 'removed'], ['took away', 'removed'], ['picked up', 'collected'], ['cleaned up', 'cleaned'], ['fixed up', 'repaired'],
    ['made sure that', 'ensured that'], ['made sure', 'ensured'], ['checked out', 'inspected'], ['told the supervisor', 'informed the supervisor'],
    ['a lot of', 'many'], ['lots of', 'many'], ['very dangerous', 'hazardous'], ['bad condition', 'unsafe condition'], ['not safe', 'unsafe'],
    ['could have hurt someone', 'could have caused an injury'], ['nobody', 'no one'], ['stuff', 'materials'], ['right away', 'immediately'], ['straight away', 'immediately'], ['at once', 'immediately'],
    ['in order to', 'to'], ['due to the fact that', 'because'], ['prior to', 'before'], ['in the event that', 'if'], ['a number of', 'several'], ['kept on', 'continued']
  ]);
  /* ---- Bahasa Indonesia and Malay shared informal words ---- */
  var IDMS = compile([
    ['yg', 'yang'], ['dgn', 'dengan'], ['dg', 'dengan'], ['utk', 'untuk'], ['untk', 'untuk'], ['dr', 'dari'], ['pd', 'pada'], ['spy', 'supaya'], ['sbg', 'sebagai'], ['tsb', 'tersebut'],
    ['dll', 'dan lain-lain'], ['dsb', 'dan sebagainya'], ['dst', 'dan seterusnya'], ['sblm', 'sebelum'], ['stlh', 'setelah'], ['bhw', 'bahawa'], ['thd', 'terhadap'], ['jg', 'juga'], ['sy', 'saya'],
    ['dilokasi', 'di lokasi'], ['dilapangan', 'di lapangan'], ['ditempat', 'di tempat'], ['diarea', 'di area'], ['disekitar', 'di sekitar'], ['dibawah', 'di bawah'], ['diatas', 'di atas'],
    ['didalam', 'di dalam'], ['diluar', 'di luar'], ['disamping', 'di samping'], ['dipinggir', 'di pinggir'], ['dijalan', 'di jalan'], ['digudang', 'di gudang'], ['dibengkel', 'di bengkel'],
    ['dipabrik', 'di pabrik'], ['dikantor', 'di kantor'], ['dilantai', 'di lantai'], ['disini', 'di sini'], ['disitu', 'di situ'], ['disana', 'di sana'], ['kesana', 'ke sana'], ['kesini', 'ke sini']
  ]);
  var ID = compile([
    ['tdk', 'tidak'], ['tdak', 'tidak'], ['gak', 'tidak'], ['ga', 'tidak'], ['nggak', 'tidak'], ['ngga', 'tidak'], ['enggak', 'tidak'], ['sdh', 'sudah'], ['udh', 'sudah'], ['udah', 'sudah'], ['sudh', 'sudah'],
    ['blm', 'belum'], ['belom', 'belum'], ['krn', 'karena'], ['krna', 'karena'], ['karna', 'karena'], ['hrs', 'harus'], ['bs', 'bisa'], ['lg', 'lagi'], ['trs', 'terus'], ['kmrn', 'kemarin'],
    ['skrg', 'sekarang'], ['msh', 'masih'], ['mgkn', 'mungkin'], ['sampe', 'sampai'], ['smpe', 'sampai'], ['kalo', 'kalau'], ['klo', 'kalau'], ['gimana', 'bagaimana'], ['gmn', 'bagaimana'],
    ['emang', 'memang'], ['banget', 'sekali'], ['bgt', 'sekali'], ['aja', 'saja'], ['doang', 'saja'], ['gitu', 'begitu'], ['gini', 'begini'], ['tau', 'tahu'], ['org', 'orang'],
    ['dicek', 'diperiksa'], ['ngecek', 'memeriksa'], ['mengecek', 'memeriksa'], ['ngebersihin', 'membersihkan'], ['dibersihin', 'dibersihkan'], ['ngerapihin', 'merapikan'], ['dirapihin', 'dirapikan'],
    ['nyingkirin', 'menyingkirkan'], ['disingkirin', 'disingkirkan'], ['benerin', 'memperbaiki'], ['ngebenerin', 'memperbaiki'], ['dibenerin', 'diperbaiki'], ['ngasih tau', 'memberi tahu'],
    ['kasih tau', 'memberi tahu'], ['dikasih tau', 'diberi tahu'], ['ngasih', 'memberi'], ['dikasih', 'diberi'], ['ngelaporin', 'melaporkan'], ['dilaporin', 'dilaporkan'],
    ['pake', 'pakai'], ['dipake', 'dipakai'], ['bikin', 'membuat'], ['dibikin', 'dibuat'], ['ngeliat', 'melihat'], ['liat', 'lihat'], ['keliatan', 'terlihat'], ['kumpulin', 'mengumpulkan'],
    ['dikumpulin', 'dikumpulkan'], ['ngumpulin', 'mengumpulkan'], ['ngumpulkan', 'mengumpulkan'], ['ngeberesin', 'membereskan'], ['beresin', 'membereskan'], ['ngangkut', 'mengangkut'], ['ngelepas', 'melepas'], ['nyalain', 'menyalakan'], ['matiin', 'mematikan'], ['ngasih tahu', 'memberi tahu'], ['nemu', 'menemukan'], ['nemuin', 'menemukan'], ['ngelihat', 'melihat'], ['ngawasin', 'mengawasi'], ['ngingetin', 'mengingatkan'], ['negur', 'menegur'], ['nutup', 'menutup'], ['masang', 'memasang'], ['ngambil', 'mengambil'], ['ngangkat', 'mengangkat'], ['nyapu', 'menyapu']
  ]);
  var MS = compile([
    ['tak', 'tidak'], ['xde', 'tiada'], ['xda', 'tiada'], ['sbb', 'sebab'], ['mcm', 'macam'], ['jgn', 'jangan'], ['dlm', 'dalam'], ['kpd', 'kepada'], ['tp', 'tetapi'], ['dpt', 'dapat'],
    ['blh', 'boleh'], ['dah', 'sudah'], ['ni', 'ini'], ['tu', 'itu'], ['sgt', 'sangat'], ['spt', 'seperti'], ['krn', 'kerana'], ['kerna', 'kerana'], ['mmg', 'memang']
  ]);

  /* ---- language guess: en / id / ms ---- */
  function set(s) { var o = {}; s.split(' ').forEach(function (w) { o[w] = 1; }); return o; }
  var EN_W = set('the and was were of to in on at for with that this he she they after before from by is are had has it his her their not found area work scaffold crew took removed because'),
      ID_W = set('yang dan di ke dari untuk pada dengan itu ini tidak sudah karena saya bahaya kabel pekerja melakukan telah akan adalah oleh sebagai tersebut kru area lokasi aman jalan bisa sebelum setelah sehingga yg dgn tdk udah gak sy'),
      MS_W = set('dan yang di untuk dengan tidak tiada telah kerana boleh sahaja mereka dalam atau pekerja tapak selamat bahaya semasa kemudian sebelum selepas supaya dapat adalah oleh daripada kepada'),
      ID_ONLY = set('karena bisa saja sehingga tersebut kru setelah gak sy udah yg bgt banget'), MS_ONLY = set('kerana boleh sahaja tiada daripada kepada selepas supaya semasa tapak');
  function detect(text) {
    var toks = String(text || '').toLowerCase().match(/[\p{L}']+/gu) || [], e = 0, i = 0, m = 0;
    toks.forEach(function (w) { if (EN_W[w]) e++; if (ID_W[w]) i += ID_ONLY[w] ? 2 : 1; if (MS_W[w]) m += MS_ONLY[w] ? 2 : 1; });
    if (Math.max(e, i, m) < 1) return 'en';
    return e >= i && e >= m ? 'en' : (i >= m ? 'id' : 'ms');
  }

  /* ---- engine ---- */
  function apply(t, c, st) {
    return t.replace(c.re, function (m, pre, hit) {
      var r = c.map[hit.toLowerCase()];
      if (r === undefined || r === hit) return m;
      if (r.charAt(0) === r.charAt(0).toLowerCase() && hit.charAt(0) !== hit.charAt(0).toLowerCase() && hit !== hit.toUpperCase()) r = r.charAt(0).toUpperCase() + r.slice(1);
      st.n++; return pre + r;
    });
  }
  var FW = 'the|a|an|is|are|was|were|and|to|of|in|on|at|for|with|that|it|he|she|we|they|yang|dan|di|ke|dari|untuk|pada|dengan|itu|ini|ada|sudah|telah|akan|tidak|atau';
  var DUP = new RegExp('(^|[^' + W + '])(' + FW + ')(\\s+\\2)+(?=$|[^' + W + '])', 'giu');
  var NOCAP = /(?:e\.g|i\.e|etc|dll|dsb|dst|vs|no|hrs|pk|bpk|bu|dr)\.\s+$/i;

  function lines(t) {
    var parts = String(t).split(/\r?\n+/).map(function (s) { return s.replace(/^\s*(?:[-•*–]|\d{1,2}[.)])\s+/, '').trim(); }).filter(Boolean);
    if (parts.length > 1) parts = parts.map(function (s) { return /[.!?]$/.test(s) ? s : s + '.'; });
    return parts.join(' ');
  }
  function fix(text, lang) {
    var orig = String(text == null ? '' : text), st = { n: 0 };
    if (!orig.trim()) return { text: '', count: 0, lang: lang || 'en' };
    lang = lang || detect(orig);
    var t = lines(orig);
    function step(fn) { var o = fn(t); if (o !== t) st.n++; t = o; }
    step(function (x) { return x.replace(/[ \t ]+/g, ' ').replace(/\s+([,.;:!?])/g, '$1').replace(/\(\s+/g, '(').replace(/\s+\)/g, ')'); });
    t = apply(t, SITE, st);
    if (lang === 'en') { t = apply(t, EN, st); t = apply(t, EN_PHR, st); t = apply(t, EN_PHR, st); }
    else {
      t = apply(t, IDMS, st); t = apply(t, lang === 'id' ? ID : MS, st);
      step(function (x) { return x.replace(/(\p{L}{3,})2(?=$|[^\p{L}\p{N}])/gu, '$1-$1').replace(/\s+nya(?=$|[^\p{L}\p{N}])/giu, 'nya'); });
    }
    t = apply(t, ACR, st);
    t = t.replace(DUP, function (m, pre, w) { st.n++; return pre + w; });
    step(function (x) { return x.replace(/([,;:])(?=\p{L})/gu, '$1 ').replace(/([\p{Ll}\p{N}][.!?])(?=\p{Lu})/gu, '$1 ').replace(/([.!?])\1+/g, '$1').replace(/,{2,}/g, ',').replace(/\s{2,}/g, ' ').trim(); });
    if (lang === 'en') t = t.replace(/(^|[^\p{L}\p{N}'’])i(?=$|[^\p{L}\p{N}])/gu, function (m, pre) { st.n++; return pre + 'I'; });
    t = t.replace(/(^|[.!?]\s+)(\p{Ll})/gu, function (m, pre, c, off, str) {
      if (pre && NOCAP.test(str.slice(0, off + pre.length))) return m;
      st.n++; return pre + c.toUpperCase();
    });
    step(function (x) { x = x.replace(/[\s,;:\-–]+$/, ''); return x && !/[.!?]$/.test(x) ? x + '.' : x; });
    if (t !== orig.trim() && st.n === 0) st.n = 1;
    return { text: t, count: t === orig.trim() ? 0 : st.n, lang: lang };
  }
  function fixAll(obj) {
    var keys = Object.keys(obj), lang = detect(keys.map(function (k) { return obj[k]; }).join(' ')), out = {}, n = 0;
    keys.forEach(function (k) { var r = fix(obj[k], lang); out[k] = r.text; n += r.count; });
    return { text: out, count: n, lang: lang };
  }
  var api = { fix: fix, fixAll: fixAll, detect: detect, version: '1.0.0' };
  root.TextFix = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
