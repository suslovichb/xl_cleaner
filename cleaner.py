from tkinter import Tk, Text, StringVar, BooleanVar, _setit, messagebox, filedialog
from tkinter.ttk import Style, Label, Button, OptionMenu, Checkbutton, Progressbar
from openpyxl import load_workbook, Workbook
import re
from os import path, system

REPLACEMENTS_PATH = "replacements.txt"

codes_dict = {
'&#09;': ' ',
'&#10;': ' ',
'&#13;': ' ',
'&#160;': ' ',
'&nbsp;': ' ',
'&#32;': ' ',
'&ndash;': '-',
'&#33;': '!',
'&#34;': '"',
'&quot;': '"',
'&#35;': '#',
'&#36;': '$',
'&#37;': '%',
'&amp;amp;': '&',
'&#38;': '&',
'&amp;': '&',
'&#39;': "'",
'&apos;': "'",
'&#40;': '(',
'&#41;': ')',
'&#42;': '*',
'&#43;': '+',
'&#44;': ',',
'&#45;': '-',
'&#46;': '.',
'&#47;': '/',
'&#48;': '0',
'&#49;': '1',
'&#50;': '2',
'&#51;': '3',
'&#52;': '4',
'&#53;': '5',
'&#54;': '6',
'&#55;': '7',
'&#56;': '8',
'&#57;': '9',
'&#58;': ':',
'&#59;': ';',
'&#60;': '<',
'&lt;': '<',
'&#61;': '=',
'&#62;': '>',
'&gt;': '>',
'&#63;': '?',
'&#64;': '@',
'&#65;': 'A',
'&#66;': 'B',
'&#67;': 'C',
'&#68;': 'D',
'&#69;': 'E',
'&#70;': 'F',
'&#71;': 'G',
'&#72;': 'H',
'&#73;': 'I',
'&#74;': 'J',
'&#75;': 'K',
'&#76;': 'L',
'&#77;': 'M',
'&#78;': 'N',
'&#79;': 'O',
'&#80;': 'P',
'&#81;': 'Q',
'&#82;': 'R',
'&#83;': 'S',
'&#84;': 'T',
'&#85;': 'U',
'&#86;': 'V',
'&#87;': 'W',
'&#88;': 'X',
'&#89;': 'Y',
'&#90;': 'Z',
'&#91;': '[',
'&#92;': '\\',
'&#93;': ']',
'&#94;': '^',
'&#95;': '_',
'&#96;': '`',
'&#97;': 'a',
'&#98;': 'b',
'&#99;': 'c',
'&#100;': 'd',
'&#101;': 'e',
'&#102;': 'f',
'&#103;': 'g',
'&#104;': 'h',
'&#105;': 'i',
'&#106;': 'j',
'&#107;': 'k',
'&#108;': 'l',
'&#109;': 'm',
'&#110;': 'n',
'&#111;': 'o',
'&#112;': 'p',
'&#113;': 'q',
'&#114;': 'r',
'&#115;': 's',
'&#116;': 't',
'&#117;': 'u',
'&#118;': 'v',
'&#119;': 'w',
'&#120;': 'x',
'&#121;': 'y',
'&#122;': 'z',
'&#123;': '{',
'&#124;': '|',
'&#125;': '}',
'&#126;': '~',
'&#161;': '¡',
'&iexcl;': '¡',
'&#162;': '¢',
'&cent;': '¢',
'&#163;': '£',
'&pound;': '£',
'&#164;': '¤',
'&curren;': '¤',
'&#165;': '¥',
'&yen;': '¥',
'&#166;': '¦',
'&brvbar;': '¦',
'&#167;': '§',
'&sect;': '§',
'&#168;': '¨',
'&uml;': '¨',
'&#169;': '©',
'&copy;': '©',
'&#170;': 'ª',
'&ordf;': 'ª',
'&#171;': '«',
'&laquo;': '«',
'&#172;': '¬',
'&not;': '¬',
'&#173;': '­',
'&shy;': '­',
'&#175;': '¯',
'&macr;': '¯',
'&#176;': '°',
'&deg;': '°',
'&#177;': '±',
'&plusmn;': '±',
'&#178;': '²',
'&sup2;': '²',
'&#179;': '³',
'&sup3;': '³',
'&#180;': '´',
'&acute;': '´',
'&#181;': 'µ',
'&micro;': 'µ',
'&#182;': '¶',
'&para;': '¶',
'&#183;': '·',
'&middot;': '·',
'&#184;': '¸',
'&cedil;': '¸',
'&#185;': '¹',
'&sup1;': '¹',
'&#186;': 'º',
'&ordm;': 'º',
'&#187;': '»',
'&raquo;': '»',
'&#188;': '¼',
'&frac14;': '¼',
'&#189;': '½',
'&frac12;': '½',
'&#190;': '¾',
'&frac34;': '¾',
'&#191;': '¿',
'&iquest;': '¿',
'&#192;': 'À',
'&Agrave;': 'À',
'&#193;': 'Á',
'&Aacute;': 'Á',
'&#194;': 'Â',
'&Acirc;': 'Â',
'&#195;': 'Ã',
'&Atilde;': 'Ã',
'&#196;': 'Ä',
'&Auml;': 'Ä',
'&#197;': 'Å',
'&Aring;': 'Å',
'&#198;': 'Æ',
'&AElig;': 'Æ',
'&#199;': 'Ç',
'&Ccedil;': 'Ç',
'&#200;': 'È',
'&Egrave;': 'È',
'&#201;': 'É',
'&Eacute;': 'É',
'&#202;': 'Ê',
'&Ecirc;': 'Ê',
'&#203;': 'Ë',
'&Euml;': 'Ë',
'&#204;': 'Ì',
'&Igrave;': 'Ì',
'&#205;': 'Í',
'&Iacute;': 'Í',
'&#206;': 'Î',
'&Icirc;': 'Î',
'&#207;': 'Ï',
'&Iuml;': 'Ï',
'&#208;': 'Ð',
'&ETH;': 'Ð',
'&#209;': 'Ñ',
'&Ntilde;': 'Ñ',
'&#210;': 'Ò',
'&Ograve;': 'Ò',
'&#211;': 'Ó',
'&Oacute;': 'Ó',
'&#212;': 'Ô',
'&Ocirc;': 'Ô',
'&#213;': 'Õ',
'&Otilde;': 'Õ',
'&#214;': 'Ö',
'&Ouml;': 'Ö',
'&#215;': '×',
'&times;': '×',
'&#216;': 'Ø',
'&Oslash;': 'Ø',
'&#217;': 'Ù',
'&Ugrave;': 'Ù',
'&#218;': 'Ú',
'&Uacute;': 'Ú',
'&#219;': 'Û',
'&Ucirc;': 'Û',
'&#220;': 'Ü',
'&Uuml;': 'Ü',
'&#221;': 'Ý',
'&Yacute;': 'Ý',
'&#222;': 'Þ',
'&THORN;': 'Þ',
'&#223;': 'ß',
'&szlig;': 'ß',
'&#224;': 'à',
'&agrave;': 'à',
'&#225;': 'á',
'&aacute;': 'á',
'&#226;': 'â',
'&;': 'â',
'&#227;': 'ã',
'&atilde;': 'ã',
'&#228;': 'ä',
'&auml;': 'ä',
'&#229;': 'å',
'&aring;': 'å',
'&#230;': 'æ',
'&aelig;': 'æ',
'&#231;': 'ç',
'&ccedil;': 'ç',
'&#232;': 'è',
'&egrave;': 'è',
'&#233;': 'é',
'&eacute;': 'é',
'&#234;': 'ê',
'&ecirc;': 'ê',
'&#235;': 'ë',
'&euml;': 'ë',
'&#236;': 'ì',
'&igrave;': 'ì',
'&#237;': 'í',
'&iacute;': 'í',
'&#238;': 'î',
'&icirc;': 'î',
'&#239;': 'ï',
'&iuml;': 'ï',
'&#240;': 'ð',
'&eth;': 'ð',
'&#241;': 'ñ',
'&ntilde;': 'ñ',
'&#242;': 'ò',
'&ograve;': 'ò',
'&#243;': 'ó',
'&oacute;': 'ó',
'&#244;': 'ô',
'&ocirc;': 'ô',
'&#245;': 'õ',
'&otilde;': 'õ',
'&#246;': 'ö',
'&ouml;': 'ö',
'&#247;': '÷',
'&divide;': '÷',
'&#248;': 'ø',
'&oslash;': 'ø',
'&#249;': 'ù',
'&ugrave;': 'ù',
'&#250;': 'ú',
'&uacute;': 'ú',
'&#251;': 'û',
'&ucirc;': 'û',
'&#252;': 'ü',
'&uuml;': 'ü',
'&#253;': 'ý',
'&yacute;': 'ý',
'&#254;': 'þ',
'&thorn;': 'þ',
'&#255;': 'ÿ',
'&yuml;': 'ÿ',
'&#38;': '&',
'&amp;': '&',
'&#8226;': '•',
'&bull;': '•',
'&#9702;': '◦',
'&#8729;': '∙',
'&#8227;': '‣',
'&#8259;': '⁃',
'&#176;': '°',
'&deg;': '°',
'&#8734;': '∞',
'&infin;': '∞',
'&#8240;': '‰',
'&permil;': '‰',
'&#8901;': '⋅',
'&sdot;': '⋅',
'&#177;': '±',
'&plusmn;': '±',
'&#8224;': '†',
'&dagger;': '†',
'&#8212;': '—',
'&mdash;': '—',
'&#172;': '¬',
'&not;': '¬',
'&#181;': 'µ',
'&micro;': 'µ',
'&#8869;': '⊥',
'&perp;': '⊥',
'&#8741;': '∥',
'&par;': '∥',
'&#36;': '$',
'&#8364;': '€',
'&euro;': '€',
'&#163;': '£',
'&pound;': '£',
'&#165;': '¥',
'&yen;': '¥',
'&#162;': '¢',
'&cent;': '¢',
'&#8377;': '₹',
'&#8360;': '₨',
'&#8369;': '₱',
'&#8361;': '₩',
'&#3647;': '฿',
'&#8363;': '₫',
'&#8362;': '₪',
'&#169;': '©',
'&copy;': '©',
'&#174;': '®',
'&reg;':'®',
'&#8471;': '℗',
'&#8482;': '™',
'&trade;': '™',
'&#8480;': '℠',
'&#945;': 'α',
'&alpha;': 'α',
'&#946;': 'β',
'&beta;': 'β',
'&#947;': 'γ',
'&gamma;': 'γ',
'&#948;': 'δ',
'&delta;': 'δ',
'&#949;': 'ε',
'&epsilon;': 'ε',
'&#950;': 'ζ',
'&zeta;': 'ζ',
'&#951;': 'η',
'&eta;': 'η',
'&#952;': 'θ',
'&theta;': 'θ',
'&#953;': 'ι',
'&iota;': 'ι',
'&#954;': 'κ',
'&kappa;': 'κ',
'&#955;': 'λ',
'&lambda;': 'λ',
'&#956;': 'μ',
'&mu;': 'μ',
'&#957;': 'ν',
'&nu;': 'ν',
'&#958;': 'ξ',
'&xi;': 'ξ',
'&#959;': 'ο',
'&omicron;': 'ο',
'&#960;': 'π',
'&pi;': 'π',
'&#961;': 'ρ',
'&rho;': 'ρ',
'&#963;': 'σ',
'&sigma;': 'σ',
'&#964;': 'τ',
'&tau;': 'τ',
'&#965;': 'υ',
'&upsilon;': 'υ',
'&#966;': 'φ',
'&phi;': 'φ',
'&#967;': 'χ',
'&chi;': 'χ',
'&#968;': 'ψ',
'&psi;': 'ψ',
'&#969;': 'ω',
'&omega;': 'ω',
'&#913;': 'Α',
'&Alpha;': 'Α',
'&#914;': 'Β',
'&Beta;': 'Β',
'&#915;': 'Γ',
'&Gamma;': 'Γ',
'&#916;': 'Δ',
'&Delta;': 'Δ',
'&#917;': 'Ε',
'&Epsilon;': 'Ε',
'&#918;': 'Ζ',
'&Zeta;': 'Ζ',
'&#919;': 'Η',
'&Eta;': 'Η',
'&#920;': 'Θ',
'&Theta;': 'Θ',
'&#921;': 'Ι',
'&Iota;': 'Ι',
'&#922;': 'Κ',
'&Kappa;': 'Κ',
'&#923;': 'Λ',
'&Lambda;': 'Λ',
'&#924;': 'Μ',
'&Mu;': 'Μ',
'&#925;': 'Ν',
'&Nu;': 'Ν',
'&#926;': 'Ξ',
'&Xi;': 'Ξ',
'&#927;': 'Ο',
'&Omicron;': 'Ο',
'&#928;': 'Π',
'&Pi;': 'Π',
'&#929;': 'Ρ',
'&Rho;': 'Ρ',
'&#931;': 'Σ',
'&Sigma;': 'Σ',
'&#932;': 'Τ',
'&Tau;': 'Τ',
'&#933;': 'Υ',
'&Upsilon;': 'Υ',
'&#934;': 'Φ',
'&Phi;': 'Φ',
'&#935;': 'Χ',
'&Chi;': 'Χ',
'&#936;': 'Ψ',
'&Psi;': 'Ψ',
'&#937;': 'Ω',
'&Omega;': 'Ω',
'&#09': ' ',
# '&#10': ' ',
'&#13': ' ',
'&#160': ' ',
'&nbsp': ' ',
'&#32': ' ',
'&#33': '!',
'&#34': '"',
'&quot': '"',
'&#35': '#',
# '&#36': '$',
'&#37': '%',
'&#38': '&',
'&amp': '&',
'&#39': "'",
'&apos': "'",
'&#40': '(',
'&#41': ')',
'&#42': '*',
'&#43': '+',
'&#44': ',',
'&#45': '-',
'&#46': '.',
'&#47': '/',
'&#48': '0',
'&#49': '1',
'&#50': '2',
'&#51': '3',
'&#52': '4',
'&#53': '5',
'&#54': '6',
'&#55': '7',
'&#56': '8',
'&#57': '9',
'&#58': ':',
'&#59': ';',
'&#60': '<',
'&lt': '<',
'&#61': '=',
'&#62': '>',
'&gt': '>',
'&#63': '?',
'&#64': '@',
'&#65': 'A',
'&#66': 'B',
'&#67': 'C',
'&#68': 'D',
'&#69': 'E',
'&#70': 'F',
'&#71': 'G',
'&#72': 'H',
'&#73': 'I',
'&#74': 'J',
'&#75': 'K',
'&#76': 'L',
'&#77': 'M',
'&#78': 'N',
'&#79': 'O',
'&#80': 'P',
'&#81': 'Q',
# '&#82': 'R',
# '&#83': 'S',
# '&#84': 'T',
'&#85': 'U',
'&#86': 'V',
# '&#87': 'W',
# '&#88': 'X',
# '&#89': 'Y',
'&#90': 'Z',
# '&#91': '[',
# '&#92': '\\',
# '&#93': ']',
# '&#94': '^',
# '&#95': '_',
# '&#96': '`',
# '&#97': 'a',
'&#98': 'b',
'&#99': 'c',
'&#100': 'd',
'&#101': 'e',
'&#102': 'f',
'&#103': 'g',
'&#104': 'h',
'&#105': 'i',
'&#106': 'j',
'&#107': 'k',
'&#108': 'l',
'&#109': 'm',
'&#110': 'n',
'&#111': 'o',
'&#112': 'p',
'&#113': 'q',
'&#114': 'r',
'&#115': 's',
'&#116': 't',
'&#117': 'u',
'&#118': 'v',
'&#119': 'w',
'&#120': 'x',
'&#121': 'y',
'&#122': 'z',
'&#123': '{',
'&#124': '|',
'&#125': '}',
'&#126': '~',
'&#161': '¡',
'&iexcl': '¡',
'&#162': '¢',
'&cent': '¢',
'&#163': '£',
'&pound': '£',
'&#164': '¤',
'&curren': '¤',
'&#165': '¥',
'&yen': '¥',
'&#166': '¦',
'&brvbar': '¦',
'&#167': '§',
'&sect': '§',
'&#168': '¨',
'&uml': '¨',
'&#169': '©',
'&copy': '©',
'&#170': 'ª',
'&ordf': 'ª',
'&#171': '«',
'&laquo': '«',
'&#172': '¬',
'&not': '¬',
'&#173': '­',
'&shy': '­',
'&#174': '®',
'&reg': '®',
'&#175': '¯',
'&macr': '¯',
'&#176': '°',
'&deg': '°',
'&#177': '±',
'&plusmn': '±',
'&#178': '²',
'&sup2': '²',
'&#179': '³',
'&sup3': '³',
'&#180': '´',
'&acute': '´',
'&#181': 'µ',
'&micro': 'µ',
'&#182': '¶',
'&para': '¶',
'&#183': '·',
'&middot': '·',
'&#184': '¸',
'&cedil': '¸',
'&#185': '¹',
'&sup1': '¹',
'&#186': 'º',
'&ordm': 'º',
'&#187': '»',
'&raquo': '»',
'&#188': '¼',
'&frac14': '¼',
'&#189': '½',
'&frac12': '½',
'&#190': '¾',
'&frac34': '¾',
'&#191': '¿',
'&iquest': '¿',
'&#192': 'À',
'&Agrave': 'À',
'&#193': 'Á',
'&Aacute': 'Á',
'&#194': 'Â',
'&Acirc': 'Â',
'&#195': 'Ã',
'&Atilde': 'Ã',
'&#196': 'Ä',
'&Auml': 'Ä',
'&#197': 'Å',
'&Aring': 'Å',
'&#198': 'Æ',
'&AElig': 'Æ',
'&#199': 'Ç',
'&Ccedil': 'Ç',
'&#200': 'È',
'&Egrave': 'È',
'&#201': 'É',
'&Eacute': 'É',
'&#202': 'Ê',
'&Ecirc': 'Ê',
'&#203': 'Ë',
'&Euml': 'Ë',
'&#204': 'Ì',
'&Igrave': 'Ì',
'&#205': 'Í',
'&Iacute': 'Í',
'&#206': 'Î',
'&Icirc': 'Î',
'&#207': 'Ï',
'&Iuml': 'Ï',
'&#208': 'Ð',
'&ETH': 'Ð',
'&#209': 'Ñ',
'&Ntilde': 'Ñ',
'&#210': 'Ò',
'&Ograve': 'Ò',
'&#211': 'Ó',
'&Oacute': 'Ó',
'&#212': 'Ô',
'&Ocirc': 'Ô',
'&#213': 'Õ',
'&Otilde': 'Õ',
'&#214': 'Ö',
'&Ouml': 'Ö',
'&#215': '×',
'&times': '×',
'&#216': 'Ø',
'&Oslash': 'Ø',
'&#217': 'Ù',
'&Ugrave': 'Ù',
'&#218': 'Ú',
'&Uacute': 'Ú',
'&#219': 'Û',
'&Ucirc': 'Û',
'&#220': 'Ü',
'&Uuml': 'Ü',
'&#221': 'Ý',
'&Yacute': 'Ý',
'&#222': 'Þ',
'&THORN': 'Þ',
'&#223': 'ß',
'&szlig': 'ß',
'&#224': 'à',
'&agrave': 'à',
'&#225': 'á',
'&aacute': 'á',
'&#226': 'â',
# '&': 'â',
'&#227': 'ã',
'&atilde': 'ã',
'&#228': 'ä',
'&auml': 'ä',
'&#229': 'å',
'&aring': 'å',
'&#230': 'æ',
'&aelig': 'æ',
'&#231': 'ç',
'&ccedil': 'ç',
'&#232': 'è',
'&egrave': 'è',
'&#233': 'é',
'&eacute': 'é',
'&#234': 'ê',
'&ecirc': 'ê',
'&#235': 'ë',
'&euml': 'ë',
'&#236': 'ì',
'&igrave': 'ì',
'&#237': 'í',
'&iacute': 'í',
'&#238': 'î',
'&icirc': 'î',
'&#239': 'ï',
'&iuml': 'ï',
'&#240': 'ð',
'&eth': 'ð',
'&#241': 'ñ',
'&ntilde': 'ñ',
'&#242': 'ò',
'&ograve': 'ò',
'&#243': 'ó',
'&oacute': 'ó',
'&#244': 'ô',
'&ocirc': 'ô',
'&#245': 'õ',
'&otilde': 'õ',
'&#246': 'ö',
'&ouml': 'ö',
'&#247': '÷',
'&divide': '÷',
'&#248': 'ø',
'&oslash': 'ø',
'&#249': 'ù',
'&ugrave': 'ù',
'&#250': 'ú',
'&uacute': 'ú',
'&#251': 'û',
'&ucirc': 'û',
'&#252': 'ü',
'&uuml': 'ü',
'&#253': 'ý',
'&yacute': 'ý',
'&#254': 'þ',
'&thorn': 'þ',
'&#255': 'ÿ',
'&yuml': 'ÿ',
'&#38': '&',
'&amp': '&',
'&#8226': '•',
'&bull': '•',
'&#9702': '◦',
'&#8729': '∙',
'&#8227': '‣',
'&#8259': '⁃',
'&#176': '°',
'&deg': '°',
'&#8734': '∞',
'&infin': '∞',
'&#8240': '‰',
'&permil': '‰',
'&#8901': '⋅',
'&sdot': '⋅',
'&#177': '±',
'&plusmn': '±',
'&#8224': '†',
'&dagger': '†',
'&#8212': '—',
'&mdash': '—',
'&#172': '¬',
'&not': '¬',
'&#181': 'µ',
'&micro': 'µ',
'&#8869': '⊥',
'&perp': '⊥',
'&#8741': '∥',
'&par': '∥',
'&#36': '$',
'&#8364': '€',
'&euro': '€',
'&#163': '£',
'&pound': '£',
'&#165': '¥',
'&yen': '¥',
'&#162': '¢',
'&cent': '¢',
'&#8377': '₹',
'&#8360': '₨',
'&#8369': '₱',
'&#8361': '₩',
'&#3647': '฿',
'&#8363': '₫',
'&#8362': '₪',
'&#169': '©',
'&copy': '©',
'&#174': '®',
'&#8471': '℗',
'&#8482': '™',
'&trade': '™',
'&#8480': '℠',
'&#945': 'α',
'&alpha': 'α',
'&#946': 'β',
'&beta': 'β',
'&#947': 'γ',
'&gamma': 'γ',
'&#948': 'δ',
'&delta': 'δ',
'&#949': 'ε',
'&epsilon': 'ε',
'&#950': 'ζ',
'&zeta': 'ζ',
'&#951': 'η',
'&eta': 'η',
'&#952': 'θ',
'&theta': 'θ',
'&#953': 'ι',
'&iota': 'ι',
'&#954': 'κ',
'&kappa': 'κ',
'&#955': 'λ',
'&lambda': 'λ',
'&#956': 'μ',
'&mu': 'μ',
'&#957': 'ν',
'&nu': 'ν',
'&#958': 'ξ',
'&xi': 'ξ',
'&#959': 'ο',
'&omicron': 'ο',
'&#960': 'π',
'&pi': 'π',
'&#961': 'ρ',
'&rho': 'ρ',
'&#963': 'σ',
'&sigma': 'σ',
'&#964': 'τ',
'&tau': 'τ',
'&#965': 'υ',
'&upsilon': 'υ',
'&#966': 'φ',
'&phi': 'φ',
'&#967': 'χ',
'&chi': 'χ',
'&#968': 'ψ',
'&psi': 'ψ',
'&#969': 'ω',
'&omega': 'ω',
'&#913': 'Α',
'&Alpha': 'Α',
'&#914': 'Β',
'&Beta': 'Β',
'&#915': 'Γ',
'&Gamma': 'Γ',
'&#916': 'Δ',
'&Delta': 'Δ',
'&#917': 'Ε',
'&Epsilon': 'Ε',
'&#918': 'Ζ',
'&Zeta': 'Ζ',
'&#919': 'Η',
'&Eta': 'Η',
'&#920': 'Θ',
'&Theta': 'Θ',
'&#921': 'Ι',
'&Iota': 'Ι',
'&#922': 'Κ',
'&Kappa': 'Κ',
'&#923': 'Λ',
'&Lambda': 'Λ',
'&#924': 'Μ',
'&Mu': 'Μ',
'&#925': 'Ν',
'&Nu': 'Ν',
'&#926': 'Ξ',
'&Xi': 'Ξ',
'&#927': 'Ο',
'&Omicron': 'Ο',
'&#928': 'Π',
'&Pi': 'Π',
'&#929': 'Ρ',
'&Rho': 'Ρ',
'&#931': 'Σ',
'&Sigma': 'Σ',
'&#932': 'Τ',
'&Tau': 'Τ',
'&#933': 'Υ',
'&Upsilon': 'Υ',
'&#934': 'Φ',
'&Phi': 'Φ',
'&#935': 'Χ',
'&Chi': 'Χ',
'&#936': 'Ψ',
'&Psi': 'Ψ',
'&#937': 'Ω',
'&Omega': 'Ω',
# '&#117;': 'u',
'&ldquo;': '"',
'&ldquo': '"',
'&rdquo;': '"',
'&rdquo': '"',
'&minus;': '-',
'&minus': '-',
'&helip;': '…',
'&helip': '…',
'&rsquo;': "'",
'&rsquo': "'",
'&lsquo;': "'",
'&lsquo': "'",
'&ndash;': '–',
'&ndash': '–'
}

sku_letters_dict = {
"," : "//0044",
'A' : '//0065',
'B' : '//0066',
'C' : '//0067',
'D' : '//0068',
'E' : '//0069',
'F' : '//0070',
'G' : '//0071',
'H' : '//0072',
'I' : '//0073',
'J' : '//0074',
'K' : '//0075',
'L' : '//0076',
'M' : '//0077',
'N' : '//0078',
'O' : '//0079',
'P' : '//0080',
'Q' : '//0081',
'R' : '//0082',
'S' : '//0083',
'T' : '//0084',
'U' : '//0085',
'V' : '//0086',
'W' : '//0087',
'X' : '//0088',
'Y' : '//0089',
'Z' : '//0090',
'a' : '//0097',
'b' : '//0098',
'c' : '//0099',
'd' : '//0100',
'e' : '//0101',
'f' : '//0102',
'g' : '//0103',
'h' : '//0104',
'i' : '//0105',
'j' : '//0106',
'k' : '//0107',
'l' : '//0108',
'm' : '//0109',
'n' : '//0110',
'o' : '//0111',
'p' : '//0112',
'q' : '//0113',
'r' : '//0114',
's' : '//0115',
't' : '//0116',
'u' : '//0117',
'v' : '//0118',
'w' : '//0119',
'x' : '//0120',
'y' : '//0121',
'z' : '//0122'
}

def get_workbook():
    try:
        filename = workbook_field.get(1.0, 'end').replace('\n','')
        return load_workbook(filename)
    except Exception:
        messagebox.showerror(title="Ошибка", message="Ошибка загрузки. Проверьте данные")
        clean_button['state'] = 'normal'
        open_button['state'] = 'normal'
        workbook_field['state'] = 'normal'

def choose_file():
    choosed_file = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"),("All files","*.*")))
    if choosed_file:
        workbook_field.delete(1.0, 'end')
        workbook_field.insert(1.0, choosed_file)
        refresh_worksheets()

def refresh_worksheets():
    worksheet_choice.set("")
    worksheet_optionmenu['menu'].delete(0, 'end')
    wb = get_workbook()
    if not wb: return

    new_choices = wb.sheetnames
    for choice in new_choices:
        worksheet_optionmenu['menu'].add_command(label=choice, command=_setit(worksheet_choice, choice))

def worksheet_changed(*args):
    worksheet_field.delete(1.0, 'end')
    worksheet_field.insert(1.0, worksheet_choice.get())

def update_counter_text(current, total):
    counter_text.set('{}/{}'.format(current, total))

def clean():
    clean_button['state'] = 'disabled'
    open_button['state'] = 'disabled'
    workbook_field['state'] = 'disabled'
    progress['value'] = 0
    process_counter = 0
    processes_number = 0
    sheet_to_read = worksheet_field.get(1.0, 'end').replace('\n','')
    wb = get_workbook()
    if not wb: return

    try:
        ws = wb[sheet_to_read]
    except Exception:
        messagebox.showerror(title="Ошибка", message="Ошибка загрузки. Проверьте данные")
        clean_button['state'] = 'normal'
        open_button['state'] = 'normal'
        workbook_field['state'] = 'normal'
        return

    if delete_columns_check.get(): processes_number += 1
    if replace_codes_check.get(): processes_number += 1
    if delete_tags_check.get(): processes_number += 1
    if replace_custom_check.get(): processes_number += 1
    if replace_letters_check.get(): processes_number += 1
    update_counter_text(process_counter, processes_number)
    process_counter_label.grid(row=11, column=0, columnspan=5, pady=(20,5))

    if delete_columns_check.get():
        process_counter += 1
        update_counter_text(process_counter, processes_number)
        delete_columns(ws)

    if replace_codes_check.get():
        process_counter += 1
        update_counter_text(process_counter, processes_number)
        replace_codes(ws)

    if delete_tags_check.get():
        process_counter += 1
        update_counter_text(process_counter, processes_number)
        delete_tags(ws)

    if replace_custom_check.get():
        process_counter += 1
        update_counter_text(process_counter, processes_number)
        replace_custom(ws)

    if replace_letters_check.get():
        process_counter += 1
        update_counter_text(process_counter, processes_number)
        replace_sku_letters(ws)

    save_workbook(wb)

    print("Done!")
    clean_button['state'] = 'normal'
    open_button['state'] = 'normal'
    workbook_field['state'] = 'normal'
    messagebox.showinfo(title="Очистка", message="Готово!")

def find_col_index(ws, col_name):
    for cell in ws[1]:
        if cell.value == col_name:
            return cell.col_idx-1


def replace_sku_letters(ws):
    progress['maximum'] = ws.max_row - 2
    progress['value'] = 0
    col_num = find_col_index(ws, 'sku') + 1
    if not col_num:
        messagebox.showerror(title="Ошибка", message="Столбец sku не найден")
        return
    print(col_num)
    for row_num in range(2, ws.max_row):
        cell = ws.cell(row_num, col_num)
        if cell.value:
            for letter, code in sku_letters_dict.items():
                print(cell.value)
                cell.value = str(cell.value).replace(letter, code)
                print(cell.value)
        progress['value'] += 1
        progress.update()
            # cell.value = str(cell.value).replace()
        # replace_in_cell(cell, sku_letters_dict)

def delete_columns(ws):
    progress["maximum"] = ws.max_column
    progress['value'] = 0
    for index, column in reversed(list(enumerate(ws.iter_cols(values_only=True), start=1))):
        if len(list(filter(None, column)))<=1 or column[0]==None:
            ws.delete_cols(index, amount=1)
        progress['value'] += 1
        progress.update()

def replace_codes(ws):
    progress['maximum'] = ws.max_row
    progress['value'] = 0
    for row in ws.iter_rows():
        row = tuple([replace_in_cell(cell, codes_dict) for cell in row])
        progress['value'] += 1
        progress.update()

def replace_in_cell(cell, dict):
    for code, character in dict.items():
        if cell.value:
            cell.value = str(cell.value).replace(code, character)

def replace_custom(ws):
    progress['maximum'] = ws.max_row
    progress['value'] = 0
    custom_dict = {}
    try:
        f = open('replacements.txt', 'r', encoding='utf8')
        text = f.read()
        arr = text.split('\n')
        f.close()
        for item in arr:
            if ' : ' in item:
                pair = item.split(' : ')
                custom_dict[pair[0]] = pair[1]
    except FileNotFoundError:
        messagebox.showerror(title="Замена", message="Файл замен не обнаружен")
        return
    except:
        messagebox.showerror(title="Ошибка", message="Ошибка настраиваемой замены")
        return

    if custom_dict:
        print(custom_dict)
        for row in ws.iter_rows():
            row = tuple([replace_in_cell(cell, custom_dict) for cell in row])
            progress['value'] += 1
            progress.update()
    else:
        messagebox.showerror(title="Замена", message="Пользовательские значения не обнаружены")

def delete_tags(ws):
    progress['value'] = 0
    progress['maximum'] = ws.max_row
    for row in ws.iter_rows():
        row = tuple([delete_tags_in_cell(cell) for cell in row])
        progress['value'] += 1
        progress.update()

def delete_tags_in_cell(cell):
    for code, character in codes_dict.items():
        if cell.value:
            cell.value = re.sub('(?!<br>)(<.*?>)',' ',str(cell.value))

def save_workbook(wb):
    filename = workbook_field.get(1.0, 'end').replace('\n','')
    if rewrite_check.get():
        wb.save(filename)
    else:
        wb.save(filename.replace(".xl","_cleaned.xl"))

def toggle_edit_button():
    if replace_custom_check.get():
        edit_custom_button['state'] = 'normal'
    else:
        edit_custom_button['state'] = 'disabled'

def edit_custom():
    if not path.isfile(REPLACEMENTS_PATH):
        with open(REPLACEMENTS_PATH, 'w'): pass
    system('start ' + REPLACEMENTS_PATH)

window = Tk()
window.title('XL Cleaner')
window.geometry('500x410')
window.resizable(0, 0)

secondary_button_style = Style()
secondary_button_style.configure('secondary.TButton', font=('Calibri',8))

workbook_label = Label(window, text="Файл")
workbook_label.grid(row=0, column=0, columnspan=5, pady=(20,5))

workbook_field = Text(window, width=50, height = 1)
workbook_field.grid(row=1, column=0, columnspan=3, padx=(10,5))

open_button = Button(window, text="Открыть", command=choose_file, style='secondary.TButton')
open_button.grid(row=1, column=4)

worksheet_label = Label(window, text="Лист")
worksheet_label.grid(row=2, column=0, columnspan=5, pady=(10,5))

worksheet_field = Text(window, width=50, height = 1)
worksheet_field.grid(row=3, column=0, columnspan=3, padx=(10,5))

worksheet_choice = StringVar()
worksheet_choice.set("")
worksheet_choice.trace("w",worksheet_changed)
worksheet_optionmenu = OptionMenu(window, worksheet_choice)
worksheet_optionmenu.grid(row=3, column=4, sticky="wns")
s = Style()
s.configure("TMenubutton", foreground="#ffffff00", background="gray83")

delete_columns_check = BooleanVar()
delete_columns_check.set(False)
delete_columns_checkbutton = Checkbutton(window, variable=delete_columns_check, onvalue=True, offvalue=False, text='Удалить пустые столбцы')
delete_columns_checkbutton.grid(row=4, column=0, columnspan=5, pady=(20,5), padx=16, sticky="w")

replace_codes_check = BooleanVar()
replace_codes_check.set(False)
replace_codes_checkbutton = Checkbutton(window, variable=replace_codes_check, onvalue=True, offvalue=False, text='Заменить HTML коды символов')
replace_codes_checkbutton.grid(row=5, column=0, columnspan=5, pady=(5,5), padx=16, sticky="w")

delete_tags_check = BooleanVar()
delete_tags_check.set(False)
delete_tags_checkbutton = Checkbutton(window, variable=delete_tags_check, onvalue=True, offvalue=False, text='Удалить HTML теги')
delete_tags_checkbutton.grid(row=6, column=0, columnspan=5, pady=(5,5), padx=16, sticky="w")

replace_custom_check = BooleanVar()
replace_custom_check.set(False)
replace_custom_checkbutton = Checkbutton(window, variable=replace_custom_check, onvalue=True, offvalue=False, text='Настраиваемая замена', command=toggle_edit_button)
replace_custom_checkbutton.grid(row=7, column=0, columnspan=2, pady=(5,5), padx=16, sticky="w")

edit_custom_button = Button(window, text="Редактировать", command=edit_custom, state = 'disabled', style='secondary.TButton')
edit_custom_button.grid(row=7, column=1, columnspan=2, ipady=0, ipadx=0)

replace_letters_check = BooleanVar()
replace_letters_check.set(False)
replace_letters_checkbutton = Checkbutton(window, variable=replace_letters_check, onvalue=True, offvalue=False, text='Заменить буквы в sku')
replace_letters_checkbutton.grid(row=8, column=0, columnspan=5, pady=(5,5), padx=16, sticky="w")

rewrite_check = BooleanVar()
rewrite_check.set(False)
rewrite_checkbutton = Checkbutton(window, variable=rewrite_check, onvalue=True, offvalue=False, text='Сохранить в исходный файл')
rewrite_checkbutton.grid(row=9, column=0, columnspan=5, pady=(5,20), padx=16, sticky="w")

clean_button = Button(window, text="ЗАПУСК", command=clean)
clean_button.grid(row=10, column=0, columnspan=5)

progress = Progressbar(window, orient='horizontal', length=490, mode = 'determinate')
progress['value'] = 0
progress.grid(row=11, column=0, columnspan=5, padx=5, pady=(20,5))

counter_text = StringVar()
counter_text.set('counter')
process_counter_label = Label(window, text="", font=("Arial", 10), textvariable=counter_text)

window.mainloop()
