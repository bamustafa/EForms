(function(G, w) {
    "object" === typeof exports && "undefined" !== typeof module ? module.exports = w() : "function" === typeof define && define.amd ? define(w) : G.formula = w()
})(this, function() {
    var G = function(w) {
        var k = function() {
                var g = {};
                g.nil = Error("#NULL!");
                g.div0 = Error("#DIV/0!");
                g.value = Error("#VALUE!");
                g.ref = Error("#REF!");
                g.name = Error("#NAME?");
                g.num = Error("#NUM!");
                g.na = Error("#N/A");
                g.error = Error("#ERROR!");
                g.data = Error("#GETTING_DATA");
                return g
            }(),
            e = function() {
                var g = {
                        flattenShallow: function(a) {
                            return a && a.reduce ?
                                a.reduce(function(c, d) {
                                    var f = Array.isArray(c),
                                        h = Array.isArray(d);
                                    return f && h ? c.concat(d) : f ? (c.push(d), c) : h ? [c].concat(d) : [c, d]
                                }) : a
                        },
                        isFlat: function(a) {
                            if (!a) return !1;
                            for (var c = 0; c < a.length; ++c)
                                if (Array.isArray(a[c])) return !1;
                            return !0
                        },
                        flatten: function() {
                            for (var a = g.argsToArray.apply(null, arguments); !g.isFlat(a);) a = g.flattenShallow(a);
                            return a
                        },
                        argsToArray: function(a) {
                            var c = [];
                            g.arrayEach(a, function(d) {
                                c.push(d)
                            });
                            return c
                        },
                        numbers: function() {
                            return this.flatten.apply(null, arguments).filter(function(a) {
                                return "number" ===
                                    typeof a
                            })
                        },
                        cleanFloat: function(a) {
                            return Math.round(1E14 * a) / 1E14
                        },
                        parseBool: function(a) {
                            if ("boolean" === typeof a || a instanceof Error) return a;
                            if ("number" === typeof a) return 0 !== a;
                            if ("string" === typeof a) {
                                var c = a.toUpperCase();
                                if ("TRUE" === c) return !0;
                                if ("FALSE" === c) return !1
                            }
                            return a instanceof Date && !isNaN(a) ? !0 : k.value
                        },
                        parseNumber: function(a) {
                            return void 0 === a || "" === a ? k.value : isNaN(a) ? k.value : parseFloat(a)
                        },
                        parseNumberArray: function(a) {
                            var c;
                            if (!a || 0 === (c = a.length)) return k.value;
                            for (var d; c--;) {
                                d =
                                    g.parseNumber(a[c]);
                                if (d === k.value) return d;
                                a[c] = d
                            }
                            return a
                        },
                        parseMatrix: function(a) {
                            if (!a || 0 === a.length) return k.value;
                            for (var c, d = 0; d < a.length; d++)
                                if (c = g.parseNumberArray(a[d]), a[d] = c, c instanceof Error) return c;
                            return a
                        }
                    },
                    b = new Date(Date.UTC(1900, 0, 1));
                g.parseDate = function(a) {
                    if (!isNaN(a)) {
                        if (a instanceof Date) return new Date(a);
                        a = parseInt(a, 10);
                        return 0 > a ? k.num : 60 >= a ? new Date(b.getTime() + 864E5 * (a - 1)) : new Date(b.getTime() + 864E5 * (a - 2))
                    }
                    return "string" !== typeof a || (a = new Date(a), isNaN(a)) ? k.value :
                        a
                };
                g.parseDateArray = function(a) {
                    for (var c = a.length, d; c--;) {
                        d = this.parseDate(a[c]);
                        if (d === k.value) return d;
                        a[c] = d
                    }
                    return a
                };
                g.anyIsError = function() {
                    for (var a = arguments.length; a--;)
                        if (arguments[a] instanceof Error) return !0;
                    return !1
                };
                g.arrayValuesToNumbers = function(a) {
                    for (var c = a.length, d; c--;) d = a[c], "number" !== typeof d && (!0 === d ? a[c] = 1 : !1 === d ? a[c] = 0 : "string" === typeof d && (d = this.parseNumber(d), a[c] = d instanceof Error ? 0 : d));
                    return a
                };
                g.rest = function(a, c) {
                    return a && "function" === typeof a.slice ? a.slice(c ||
                        1) : a
                };
                g.initial = function(a, c) {
                    return a && "function" === typeof a.slice ? a.slice(0, a.length - (c || 1)) : a
                };
                g.arrayEach = function(a, c) {
                    for (var d = -1, f = a.length; ++d < f && !1 !== c(a[d], d, a););
                    return a
                };
                g.transpose = function(a) {
                    return a ? a[0].map(function(c, d) {
                        return a.map(function(f) {
                            return f[d]
                        })
                    }) : k.value
                };
                return g
            }(),
            y = {};
        y.datetime = function() {
            function g(d) {
                return (d - a) / 864E5 + (-22038912E5 < d ? 2 : 1)
            }
            var b = {},
                a = new Date(1900, 0, 1),
                c = [
                    [],
                    [1, 2, 3, 4, 5, 6, 7],
                    [7, 1, 2, 3, 4, 5, 6],
                    [6, 0, 1, 2, 3, 4, 5],
                    [],
                    [],
                    [],
                    [],
                    [],
                    [],
                    [],
                    [7, 1, 2, 3, 4, 5, 6],
                    [6, 7, 1, 2, 3, 4, 5],
                    [5, 6, 7, 1, 2, 3, 4],
                    [4, 5, 6, 7, 1, 2, 3],
                    [3, 4, 5, 6, 7, 1, 2],
                    [2, 3, 4, 5, 6, 7, 1],
                    [1, 2, 3, 4, 5, 6, 7]
                ];
            b.DATE = function(d, f, h) {
                d = e.parseNumber(d);
                f = e.parseNumber(f);
                h = e.parseNumber(h);
                return e.anyIsError(d, f, h) ? k.value : 0 > d || 0 > f || 0 > h ? k.num : new Date(d, f - 1, h)
            };
            b.DATEVALUE = function(d) {
                if ("string" !== typeof d) return k.value;
                d = Date.parse(d);
                return isNaN(d) ? k.value : -22038912E5 >= d ? (d - a) / 864E5 + 1 : (d - a) / 864E5 + 2
            };
            b.DAY = function(d) {
                d = e.parseDate(d);
                return d instanceof Error ? d : d.getDate()
            };
            b.DAYS = function(d, f) {
                d = e.parseDate(d);
                f = e.parseDate(f);
                return d instanceof Error ? d : f instanceof Error ? f : g(d) - g(f)
            };
            b.DAYS360 = function(d, f, h) {};
            b.EDATE = function(d, f) {
                d = e.parseDate(d);
                if (d instanceof Error) return d;
                if (isNaN(f)) return k.value;
                f = parseInt(f, 10);
                d.setMonth(d.getMonth() + f);
                return g(d)
            };
            b.EOMONTH = function(d, f) {
                d = e.parseDate(d);
                if (d instanceof Error) return d;
                if (isNaN(f)) return k.value;
                f = parseInt(f, 10);
                return g(new Date(d.getFullYear(), d.getMonth() + f + 1, 0))
            };
            b.HOUR = function(d) {
                d = e.parseDate(d);
                return d instanceof Error ? d : d.getHours()
            };
            b.INTERVAL = function(d) {
                if ("number" !== typeof d && "string" !== typeof d) return k.value;
                d = parseInt(d, 10);
                var f = Math.floor(d / 94608E4);
                d %= 94608E4;
                var h = Math.floor(d / 2592E3);
                d %= 2592E3;
                var l = Math.floor(d / 86400);
                d %= 86400;
                var m = Math.floor(d / 3600);
                d %= 3600;
                var n = Math.floor(d / 60);
                d %= 60;
                return "P" + (0 < f ? f + "Y" : "") + (0 < h ? h + "M" : "") + (0 < l ? l + "D" : "") + "T" + (0 < m ? m + "H" : "") + (0 < n ? n + "M" : "") + (0 < d ? d + "S" : "")
            };
            b.ISOWEEKNUM = function(d) {
                d = e.parseDate(d);
                if (d instanceof Error) return d;
                d.setHours(0, 0, 0);
                d.setDate(d.getDate() + 4 - (d.getDay() ||
                    7));
                var f = new Date(d.getFullYear(), 0, 1);
                return Math.ceil(((d - f) / 864E5 + 1) / 7)
            };
            b.MINUTE = function(d) {
                d = e.parseDate(d);
                return d instanceof Error ? d : d.getMinutes()
            };
            b.MONTH = function(d) {
                d = e.parseDate(d);
                return d instanceof Error ? d : d.getMonth() + 1
            };
            b.NETWORKDAYS = function(d, f, h) {};
            b.NETWORKDAYS.INTL = function(d, f, h, l) {};
            b.NOW = function() {
                return new Date
            };
            b.SECOND = function(d) {
                d = e.parseDate(d);
                return d instanceof Error ? d : d.getSeconds()
            };
            b.TIME = function(d, f, h) {
                d = e.parseNumber(d);
                f = e.parseNumber(f);
                h = e.parseNumber(h);
                return e.anyIsError(d, f, h) ? k.value : 0 > d || 0 > f || 0 > h ? k.num : (3600 * d + 60 * f + h) / 86400
            };
            b.TIMEVALUE = function(d) {
                d = e.parseDate(d);
                return d instanceof Error ? d : (3600 * d.getHours() + 60 * d.getMinutes() + d.getSeconds()) / 86400
            };
            b.TODAY = function() {
                return new Date
            };
            b.WEEKDAY = function(d, f) {
                d = e.parseDate(d);
                if (d instanceof Error) return d;
                void 0 === f && (f = 1);
                d = d.getDay();
                return c[f][d]
            };
            b.WEEKNUM = function(d, f) {};
            b.WORKDAY = function(d, f, h) {};
            b.WORKDAY.INTL = function(d, f, h, l) {};
            b.YEAR = function(d) {
                d = e.parseDate(d);
                return d instanceof
                Error ? d : d.getFullYear()
            };
            b.YEARFRAC = function(d, f, h) {};
            return b
        }();
        y.database = function() {
            function g(a, c) {
                for (var d = {}, f = 1; f < a[0].length; ++f) d[f] = !0;
                var h = c[0].length;
                for (f = 1; f < c.length; ++f) c[f].length > h && (h = c[f].length);
                for (f = 1; f < a.length; ++f)
                    for (var l = 1; l < a[f].length; ++l) {
                        for (var m = !1, n = !1, p = 0; p < c.length; ++p) {
                            var q = c[p];
                            if (!(q.length < h) && a[f][0] === q[0]) {
                                n = !0;
                                for (var t = 1; t < q.length; ++t) m = m || eval(a[f][l] + q[t])
                            }
                        }
                        n && (d[l] = d[l] && m)
                    }
                c = [];
                for (h = 0; h < a[0].length; ++h) d[h] && c.push(h - 1);
                return c
            }
            var b = {
                FINDFIELD: function(a,
                    c) {
                    for (var d = null, f = 0; f < a.length; f++)
                        if (a[f][0] === c) {
                            d = f;
                            break
                        } return null == d ? k.value : d
                },
                DAVERAGE: function(a, c, d) {
                    if (isNaN(c) && "string" !== typeof c) return k.value;
                    d = g(a, d);
                    "string" === typeof c ? (c = b.FINDFIELD(a, c), a = e.rest(a[c])) : a = e.rest(a[c]);
                    for (var f = c = 0; f < d.length; f++) c += a[d[f]];
                    return 0 === d.length ? k.div0 : c / d.length
                },
                DCOUNT: function(a, c, d) {},
                DCOUNTA: function(a, c, d) {},
                DGET: function(a, c, d) {
                    if (isNaN(c) && "string" !== typeof c) return k.value;
                    d = g(a, d);
                    "string" === typeof c ? (c = b.FINDFIELD(a, c), a = e.rest(a[c])) :
                        a = e.rest(a[c]);
                    return 0 === d.length ? k.value : 1 < d.length ? k.num : a[d[0]]
                },
                DMAX: function(a, c, d) {
                    if (isNaN(c) && "string" !== typeof c) return k.value;
                    d = g(a, d);
                    "string" === typeof c ? (c = b.FINDFIELD(a, c), a = e.rest(a[c])) : a = e.rest(a[c]);
                    c = a[d[0]];
                    for (var f = 1; f < d.length; f++) c < a[d[f]] && (c = a[d[f]]);
                    return c
                },
                DMIN: function(a, c, d) {
                    if (isNaN(c) && "string" !== typeof c) return k.value;
                    d = g(a, d);
                    "string" === typeof c ? (c = b.FINDFIELD(a, c), a = e.rest(a[c])) : a = e.rest(a[c]);
                    c = a[d[0]];
                    for (var f = 1; f < d.length; f++) c > a[d[f]] && (c = a[d[f]]);
                    return c
                },
                DPRODUCT: function(a, c, d) {
                    if (isNaN(c) && "string" !== typeof c) return k.value;
                    d = g(a, d);
                    if ("string" === typeof c) {
                        c = b.FINDFIELD(a, c);
                        var f = e.rest(a[c])
                    } else f = e.rest(a[c]);
                    a = [];
                    for (c = 0; c < d.length; c++) a[c] = f[d[c]];
                    if (d = a)
                        for (a = [], c = 0; c < d.length; ++c) d[c] && a.push(d[c]);
                    else a = d;
                    d = 1;
                    for (c = 0; c < a.length; c++) d *= a[c];
                    return d
                },
                DSTDEV: function(a, c, d) {},
                DSTDEVP: function(a, c, d) {},
                DSUM: function(a, c, d) {},
                DVAR: function(a, c, d) {},
                DVARP: function(a, c, d) {},
                MATCH: function(a, c, d) {
                    if (!a && !c) return k.na;
                    2 === arguments.length && (d =
                        1);
                    if (!(c instanceof Array) || -1 !== d && 0 !== d && 1 !== d) return k.na;
                    for (var f, h, l = 0; l < c.length; l++)
                        if (1 === d) {
                            if (c[l] === a) return l + 1;
                            c[l] < a && (h ? c[l] > h && (f = l + 1, h = c[l]) : (f = l + 1, h = c[l]))
                        } else if (0 === d)
                        if ("string" === typeof a) {
                            if (a = a.replace(/\?/g, "."), c[l].toLowerCase().match(a.toLowerCase())) return l + 1
                        } else {
                            if (c[l] === a) return l + 1
                        }
                    else if (-1 === d) {
                        if (c[l] === a) return l + 1;
                        c[l] > a && (h ? c[l] < h && (f = l + 1, h = c[l]) : (f = l + 1, h = c[l]))
                    }
                    return f ? f : k.na
                }
            };
            return b
        }();
        y.engineering = function() {
            var g = {
                BESSELI: function(b, a) {},
                BESSELJ: function(b,
                    a) {},
                BESSELK: function(b, a) {},
                BESSELY: function(b, a) {},
                BIN2DEC: function(b) {
                    if (!/^[01]{1,10}$/.test(b)) return k.num;
                    var a = parseInt(b, 2);
                    b = b.toString();
                    return 10 === b.length && "1" === b.substring(0, 1) ? parseInt(b.substring(1), 2) - 512 : a
                },
                BIN2HEX: function(b, a) {
                    if (!/^[01]{1,10}$/.test(b)) return k.num;
                    var c = b.toString();
                    if (10 === c.length && "1" === c.substring(0, 1)) return (0xfffffffe00 + parseInt(c.substring(1), 2)).toString(16);
                    b = parseInt(b, 2).toString(16);
                    if (void 0 === a) return b;
                    if (isNaN(a)) return k.value;
                    if (0 > a) return k.num;
                    a = Math.floor(a);
                    return a >= b.length ? REPT("0", a - b.length) + b : k.num
                },
                BIN2OCT: function(b, a) {
                    if (!/^[01]{1,10}$/.test(b)) return k.num;
                    var c = b.toString();
                    if (10 === c.length && "1" === c.substring(0, 1)) return (1073741312 + parseInt(c.substring(1), 2)).toString(8);
                    b = parseInt(b, 2).toString(8);
                    if (void 0 === a) return b;
                    if (isNaN(a)) return k.value;
                    if (0 > a) return k.num;
                    a = Math.floor(a);
                    return a >= b.length ? REPT("0", a - b.length) + b : k.num
                },
                BITAND: function(b, a) {
                    b = e.parseNumber(b);
                    a = e.parseNumber(a);
                    return e.anyIsError(b, a) ? k.value : 0 >
                        b || 0 > a || Math.floor(b) !== b || Math.floor(a) !== a || 0xffffffffffff < b || 0xffffffffffff < a ? k.num : b & a
                },
                BITLSHIFT: function(b, a) {
                    b = e.parseNumber(b);
                    a = e.parseNumber(a);
                    return e.anyIsError(b, a) ? k.value : 0 > b || Math.floor(b) !== b || 0xffffffffffff < b || 53 < Math.abs(a) ? k.num : 0 <= a ? b << a : b >> -a
                },
                BITOR: function(b, a) {
                    b = e.parseNumber(b);
                    a = e.parseNumber(a);
                    return e.anyIsError(b, a) ? k.value : 0 > b || 0 > a || Math.floor(b) !== b || Math.floor(a) !== a || 0xffffffffffff < b || 0xffffffffffff < a ? k.num : b | a
                },
                BITRSHIFT: function(b, a) {
                    b = e.parseNumber(b);
                    a = e.parseNumber(a);
                    return e.anyIsError(b, a) ? k.value : 0 > b || Math.floor(b) !== b || 0xffffffffffff < b || 53 < Math.abs(a) ? k.num : 0 <= a ? b >> a : b << -a
                },
                BITXOR: function(b, a) {
                    b = e.parseNumber(b);
                    a = e.parseNumber(a);
                    return e.anyIsError(b, a) ? k.value : 0 > b || 0 > a || Math.floor(b) !== b || Math.floor(a) !== a || 0xffffffffffff < b || 0xffffffffffff < a ? k.num : b ^ a
                },
                COMPLEX: function(b, a, c) {
                    b = e.parseNumber(b);
                    a = e.parseNumber(a);
                    if (e.anyIsError(b, a)) return b;
                    c = void 0 === c ? "i" : c;
                    return "i" !== c && "j" !== c ? k.value : 0 === b && 0 === a ? 0 : 0 === b ? 1 === a ? c : a.toString() + c : 0 === a ? b.toString() :
                        b.toString() + (0 < a ? "+" : "") + (1 === a ? c : a.toString() + c)
                },
                CONVERT: function(b, a, c) {
                    b = e.parseNumber(b);
                    if (b instanceof Error) return b;
                    for (var d = [
                            ["a.u. of action", "?", null, "action", !1, !1, 1.05457168181818E-34],
                            ["a.u. of charge", "e", null, "electric_charge", !1, !1, 1.60217653141414E-19],
                            ["a.u. of energy", "Eh", null, "energy", !1, !1, 4.35974417757576E-18],
                            ["a.u. of length", "a?", null, "length", !1, !1, 5.29177210818182E-11],
                            ["a.u. of mass", "m?", null, "mass", !1, !1, 9.10938261616162E-31],
                            ["a.u. of time", "?/Eh", null, "time", !1,
                                !1, 2.41888432650516E-17
                            ],
                            ["admiralty knot", "admkn", null, "speed", !1, !0, .514773333],
                            ["ampere", "A", null, "electric_current", !0, !1, 1],
                            ["ampere per meter", "A/m", null, "magnetic_field_intensity", !0, !1, 1],
                            ["\u00e5ngstr\u00f6m", "\u00c5", ["ang"], "length", !1, !0, 1E-10],
                            ["are", "ar", null, "area", !1, !0, 100],
                            ["astronomical unit", "ua", null, "length", !1, !1, 1.49597870691667E-11],
                            ["bar", "bar", null, "pressure", !1, !1, 1E5],
                            ["barn", "b", null, "area", !1, !1, 1E-28],
                            ["becquerel", "Bq", null, "radioactivity", !0, !1, 1],
                            ["bit", "bit", ["b"],
                                "information", !1, !0, 1
                            ],
                            ["btu", "BTU", ["btu"], "energy", !1, !0, 1055.05585262],
                            ["byte", "byte", null, "information", !1, !0, 8],
                            ["candela", "cd", null, "luminous_intensity", !0, !1, 1],
                            ["candela per square metre", "cd/m?", null, "luminance", !0, !1, 1],
                            ["coulomb", "C", null, "electric_charge", !0, !1, 1],
                            ["cubic \u00e5ngstr\u00f6m", "ang3", ["ang^3"], "volume", !1, !0, 1E-30],
                            ["cubic foot", "ft3", ["ft^3"], "volume", !1, !0, .028316846592],
                            ["cubic inch", "in3", ["in^3"], "volume", !1, !0, 1.6387064E-5],
                            ["cubic light-year", "ly3", ["ly^3"], "volume",
                                !1, !0, 8.46786664623715E-47
                            ],
                            ["cubic metre", "m?", null, "volume", !0, !0, 1],
                            ["cubic mile", "mi3", ["mi^3"], "volume", !1, !0, 4.16818182544058E9],
                            ["cubic nautical mile", "Nmi3", ["Nmi^3"], "volume", !1, !0, 6352182208],
                            ["cubic Pica", "Pica3", ["Picapt3", "Pica^3", "Picapt^3"], "volume", !1, !0, 7.58660370370369E-8],
                            ["cubic yard", "yd3", ["yd^3"], "volume", !1, !0, .764554857984],
                            ["cup", "cup", null, "volume", !1, !0, 2.365882365E-4],
                            ["dalton", "Da", ["u"], "mass", !1, !1, 1.66053886282828E-27],
                            ["day", "d", ["day"], "time", !1, !0, 86400],
                            ["degree",
                                "\u00b0", null, "angle", !1, !1, .0174532925199433
                            ],
                            ["degrees Rankine", "Rank", null, "temperature", !1, !0, .555555555555556],
                            ["dyne", "dyn", ["dy"], "force", !1, !0, 1E-5],
                            ["electronvolt", "eV", ["ev"], "energy", !1, !0, 1.60217656514141],
                            ["ell", "ell", null, "length", !1, !0, 1.143],
                            ["erg", "erg", ["e"], "energy", !1, !0, 1E-7],
                            ["farad", "F", null, "electric_capacitance", !0, !1, 1],
                            ["fluid ounce", "oz", null, "volume", !1, !0, 2.95735295625E-5],
                            ["foot", "ft", null, "length", !1, !0, .3048],
                            ["foot-pound", "flb", null, "energy", !1, !0, 1.3558179483314],
                            ["gal", "Gal", null, "acceleration", !1, !1, .01],
                            ["gallon", "gal", null, "volume", !1, !0, .003785411784],
                            ["gauss", "G", ["ga"], "magnetic_flux_density", !1, !0, 1],
                            ["grain", "grain", null, "mass", !1, !0, 6.47989E-5],
                            ["gram", "g", null, "mass", !1, !0, .001],
                            ["gray", "Gy", null, "absorbed_dose", !0, !1, 1],
                            ["gross registered ton", "GRT", ["regton"], "volume", !1, !0, 2.8316846592],
                            ["hectare", "ha", null, "area", !1, !0, 1E4],
                            ["henry", "H", null, "inductance", !0, !1, 1],
                            ["hertz", "Hz", null, "frequency", !0, !1, 1],
                            ["horsepower", "HP", ["h"], "power", !1, !0,
                                745.69987158227
                            ],
                            ["horsepower-hour", "HPh", ["hh", "hph"], "energy", !1, !0, 2684519.538],
                            ["hour", "h", ["hr"], "time", !1, !0, 3600],
                            ["imperial gallon (U.K.)", "uk_gal", null, "volume", !1, !0, .00454609],
                            ["imperial hundredweight", "lcwt", ["uk_cwt", "hweight"], "mass", !1, !0, 50.802345],
                            ["imperial quart (U.K)", "uk_qt", null, "volume", !1, !0, .0011365225],
                            ["imperial ton", "brton", ["uk_ton", "LTON"], "mass", !1, !0, 1016.046909],
                            ["inch", "in", null, "length", !1, !0, .0254],
                            ["international acre", "uk_acre", null, "area", !1, !0, 4046.8564224],
                            ["IT calorie", "cal", null, "energy", !1, !0, 4.1868],
                            ["joule", "J", null, "energy", !0, !0, 1],
                            ["katal", "kat", null, "catalytic_activity", !0, !1, 1],
                            ["kelvin", "K", ["kel"], "temperature", !0, !0, 1],
                            ["kilogram", "kg", null, "mass", !0, !0, 1],
                            ["knot", "kn", null, "speed", !1, !0, .514444444444444],
                            ["light-year", "ly", null, "length", !1, !0, 9460730472580800],
                            ["litre", "L", ["l", "lt"], "volume", !1, !0, .001],
                            ["lumen", "lm", null, "luminous_flux", !0, !1, 1],
                            ["lux", "lx", null, "illuminance", !0, !1, 1],
                            ["maxwell", "Mx", null, "magnetic_flux", !1, !1, 1E-18],
                            ["measurement ton", "MTON", null, "volume", !1, !0, 1.13267386368],
                            ["meter per hour", "m/h", ["m/hr"], "speed", !1, !0, 2.7777777777778E-4],
                            ["meter per second", "m/s", ["m/sec"], "speed", !0, !0, 1],
                            ["meter per second squared", "m?s??", null, "acceleration", !0, !1, 1],
                            ["parsec", "pc", ["parsec"], "length", !1, !0, 0x6da012f958ee1c],
                            ["meter squared per second", "m?/s", null, "kinematic_viscosity", !0, !1, 1],
                            ["metre", "m", null, "length", !0, !0, 1],
                            ["miles per hour", "mph", null, "speed", !1, !0, .44704],
                            ["millimetre of mercury", "mmHg", null, "pressure",
                                !1, !1, 133.322
                            ],
                            ["minute", "?", null, "angle", !1, !1, 2.90888208665722E-4],
                            ["minute", "min", ["mn"], "time", !1, !0, 60],
                            ["modern teaspoon", "tspm", null, "volume", !1, !0, 5E-6],
                            ["mole", "mol", null, "amount_of_substance", !0, !1, 1],
                            ["morgen", "Morgen", null, "area", !1, !0, 2500],
                            ["n.u. of action", "?", null, "action", !1, !1, 1.05457168181818E-34],
                            ["n.u. of mass", "m?", null, "mass", !1, !1, 9.10938261616162E-31],
                            ["n.u. of speed", "c?", null, "speed", !1, !1, 299792458],
                            ["n.u. of time", "?/(me?c??)", null, "time", !1, !1, 1.28808866778687E-21],
                            ["nautical mile",
                                "M", ["Nmi"], "length", !1, !0, 1852
                            ],
                            ["newton", "N", null, "force", !0, !0, 1],
                            ["\u0153rsted", "Oe ", null, "magnetic_field_intensity", !1, !1, 79.5774715459477],
                            ["ohm", "\u03a9", null, "electric_resistance", !0, !1, 1],
                            ["ounce mass", "ozm", null, "mass", !1, !0, .028349523125],
                            ["pascal", "Pa", null, "pressure", !0, !1, 1],
                            ["pascal second", "Pa?s", null, "dynamic_viscosity", !0, !1, 1],
                            ["pferdest\u00e4rke", "PS", null, "power", !1, !0, 735.49875],
                            ["phot", "ph", null, "illuminance", !1, !1, 1E-4],
                            ["pica (1/6 inch)", "pica", null, "length", !1, !0, 3.5277777777778E-4],
                            ["pica (1/72 inch)", "Pica", ["Picapt"], "length", !1, !0, .00423333333333333],
                            ["poise", "P", null, "dynamic_viscosity", !1, !1, .1],
                            ["pond", "pond", null, "force", !1, !0, .00980665],
                            ["pound force", "lbf", null, "force", !1, !0, 4.4482216152605],
                            ["pound mass", "lbm", null, "mass", !1, !0, .45359237],
                            ["quart", "qt", null, "volume", !1, !0, 9.46352946E-4],
                            ["radian", "rad", null, "angle", !0, !1, 1],
                            ["second", "?", null, "angle", !1, !1, 4.84813681109536E-6],
                            ["second", "s", ["sec"], "time", !0, !0, 1],
                            ["short hundredweight", "cwt", ["shweight"], "mass", !1,
                                !0, 45.359237
                            ],
                            ["siemens", "S", null, "electrical_conductance", !0, !1, 1],
                            ["sievert", "Sv", null, "equivalent_dose", !0, !1, 1],
                            ["slug", "sg", null, "mass", !1, !0, 14.59390294],
                            ["square \u00e5ngstr\u00f6m", "ang2", ["ang^2"], "area", !1, !0, 1E-20],
                            ["square foot", "ft2", ["ft^2"], "area", !1, !0, .09290304],
                            ["square inch", "in2", ["in^2"], "area", !1, !0, 6.4516E-4],
                            ["square light-year", "ly2", ["ly^2"], "area", !1, !0, 8.95054210748189E31],
                            ["square meter", "m?", null, "area", !0, !0, 1],
                            ["square mile", "mi2", ["mi^2"], "area", !1, !0, 2589988.110336],
                            ["square nautical mile", "Nmi2", ["Nmi^2"], "area", !1, !0, 3429904],
                            ["square Pica", "Pica2", ["Picapt2", "Pica^2", "Picapt^2"], "area", !1, !0, 1.792111111111E-5],
                            ["square yard", "yd2", ["yd^2"], "area", !1, !0, .83612736],
                            ["statute mile", "mi", null, "length", !1, !0, 1609.344],
                            ["steradian", "sr", null, "solid_angle", !0, !1, 1],
                            ["stilb", "sb", null, "luminance", !1, !1, 1E-4],
                            ["stokes", "St", null, "kinematic_viscosity", !1, !1, 1E-4],
                            ["stone", "stone", null, "mass", !1, !0, 6.35029318],
                            ["tablespoon", "tbs", null, "volume", !1, !0, 1.47868E-5],
                            ["teaspoon",
                                "tsp", null, "volume", !1, !0, 4.92892E-6
                            ],
                            ["tesla", "T", null, "magnetic_flux_density", !0, !0, 1],
                            ["thermodynamic calorie", "c", null, "energy", !1, !0, 4.184],
                            ["ton", "ton", null, "mass", !1, !0, 907.18474],
                            ["tonne", "t", null, "mass", !1, !1, 1E3],
                            ["U.K. pint", "uk_pt", null, "volume", !1, !0, 5.6826125E-4],
                            ["U.S. bushel", "bushel", null, "volume", !1, !0, .03523907],
                            ["U.S. oil barrel", "barrel", null, "volume", !1, !0, .158987295],
                            ["U.S. pint", "pt", ["us_pt"], "volume", !1, !0, 4.73176473E-4],
                            ["U.S. survey mile", "survey_mi", null, "length", !1, !0,
                                1609.347219
                            ],
                            ["U.S. survey/statute acre", "us_acre", null, "area", !1, !0, 4046.87261],
                            ["volt", "V", null, "voltage", !0, !1, 1],
                            ["watt", "W", null, "power", !0, !0, 1],
                            ["watt-hour", "Wh", ["wh"], "energy", !1, !0, 3600],
                            ["weber", "Wb", null, "magnetic_flux", !0, !1, 1],
                            ["yard", "yd", null, "length", !1, !0, .9144],
                            ["year", "yr", null, "time", !1, !0, 31557600]
                        ], f = {
                            Yi: ["yobi", 80, 1.2089258196146292E24, "Yi", "yotta"],
                            Zi: ["zebi", 70, 1.1805916207174113E21, "Zi", "zetta"],
                            Ei: ["exbi", 60, 0x1000000000000000, "Ei", "exa"],
                            Pi: ["pebi", 50, 0x4000000000000,
                                "Pi", "peta"
                            ],
                            Ti: ["tebi", 40, 1099511627776, "Ti", "tera"],
                            Gi: ["gibi", 30, 1073741824, "Gi", "giga"],
                            Mi: ["mebi", 20, 1048576, "Mi", "mega"],
                            ki: ["kibi", 10, 1024, "ki", "kilo"]
                        }, h = {
                            Y: ["yotta", 1E24, "Y"],
                            Z: ["zetta", 1E21, "Z"],
                            E: ["exa", 1E18, "E"],
                            P: ["peta", 1E15, "P"],
                            T: ["tera", 1E12, "T"],
                            G: ["giga", 1E9, "G"],
                            M: ["mega", 1E6, "M"],
                            k: ["kilo", 1E3, "k"],
                            h: ["hecto", 100, "h"],
                            e: ["dekao", 10, "e"],
                            d: ["deci", .1, "d"],
                            c: ["centi", .01, "c"],
                            m: ["milli", .001, "m"],
                            u: ["micro", 1E-6, "u"],
                            n: ["nano", 1E-9, "n"],
                            p: ["pico", 1E-12, "p"],
                            f: ["femto", 1E-15, "f"],
                            a: ["atto", 1E-18, "a"],
                            z: ["zepto", 1E-21, "z"],
                            y: ["yocto", 1E-24, "y"]
                        }, l = null, m = null, n = a, p = c, q = 1, t = 1, u, r = 0; r < d.length; r++) {
                        u = null === d[r][2] ? [] : d[r][2];
                        if (d[r][1] === n || 0 <= u.indexOf(n)) l = d[r];
                        if (d[r][1] === p || 0 <= u.indexOf(p)) m = d[r]
                    }
                    if (null === l)
                        for (u = f[a.substring(0, 2)], r = h[a.substring(0, 1)], "da" === a.substring(0, 2) && (r = ["dekao", 10, "da"]), u ? (q = u[2], n = a.substring(2)) : r && (q = r[1], n = a.substring(r[2].length)), a = 0; a < d.length; a++)
                            if (u = null === d[a][2] ? [] : d[a][2], d[a][1] === n || 0 <= u.indexOf(n)) l = d[a];
                    if (null === m)
                        for (f =
                            f[c.substring(0, 2)], h = h[c.substring(0, 1)], "da" === c.substring(0, 2) && (h = ["dekao", 10, "da"]), f ? (t = f[2], p = c.substring(2)) : h && (t = h[1], p = c.substring(h[2].length)), c = 0; c < d.length; c++)
                            if (u = null === d[c][2] ? [] : d[c][2], d[c][1] === p || 0 <= u.indexOf(p)) m = d[c];
                    return null === l || null === m || l[3] !== m[3] ? k.na : b * l[6] * q / (m[6] * t)
                },
                DEC2BIN: function(b, a) {
                    b = e.parseNumber(b);
                    if (b instanceof Error) return b;
                    if (!/^-?[0-9]{1,3}$/.test(b) || -512 > b || 511 < b) return k.num;
                    if (0 > b) return "1" + REPT("0", 9 - (512 + b).toString(2).length) + (512 + b).toString(2);
                    b = parseInt(b, 10).toString(2);
                    if ("undefined" === typeof a) return b;
                    if (isNaN(a)) return k.value;
                    if (0 > a) return k.num;
                    a = Math.floor(a);
                    return a >= b.length ? REPT("0", a - b.length) + b : k.num
                },
                DEC2HEX: function(b, a) {
                    b = e.parseNumber(b);
                    if (b instanceof Error) return b;
                    if (!/^-?[0-9]{1,12}$/.test(b) || -549755813888 > b || 549755813887 < b) return k.num;
                    if (0 > b) return (1099511627776 + b).toString(16);
                    b = parseInt(b, 10).toString(16);
                    if ("undefined" === typeof a) return b;
                    if (isNaN(a)) return k.value;
                    if (0 > a) return k.num;
                    a = Math.floor(a);
                    return a >=
                        b.length ? REPT("0", a - b.length) + b : k.num
                },
                DEC2OCT: function(b, a) {
                    b = e.parseNumber(b);
                    if (b instanceof Error) return b;
                    if (!/^-?[0-9]{1,9}$/.test(b) || -536870912 > b || 536870911 < b) return k.num;
                    if (0 > b) return (1073741824 + b).toString(8);
                    b = parseInt(b, 10).toString(8);
                    if ("undefined" === typeof a) return b;
                    if (isNaN(a)) return k.value;
                    if (0 > a) return k.num;
                    a = Math.floor(a);
                    return a >= b.length ? REPT("0", a - b.length) + b : k.num
                },
                DELTA: function(b, a) {
                    a = void 0 === a ? 0 : a;
                    b = e.parseNumber(b);
                    a = e.parseNumber(a);
                    return e.anyIsError(b, a) ? k.value :
                        b === a ? 1 : 0
                },
                ERF: function(b, a) {}
            };
            g.ERF.PRECISE = function() {};
            g.ERFC = function(b) {};
            g.ERFC.PRECISE = function() {};
            g.GESTEP = function(b, a) {
                a = a || 0;
                b = e.parseNumber(b);
                return e.anyIsError(a, b) ? b : b >= a ? 1 : 0
            };
            g.HEX2BIN = function(b, a) {
                if (!/^[0-9A-Fa-f]{1,10}$/.test(b)) return k.num;
                var c = 10 === b.length && "f" === b.substring(0, 1).toLowerCase() ? !0 : !1;
                b = c ? parseInt(b, 16) - 1099511627776 : parseInt(b, 16);
                if (-512 > b || 511 < b) return k.num;
                if (c) return "1" + REPT("0", 9 - (512 + b).toString(2).length) + (512 + b).toString(2);
                c = b.toString(2);
                if (void 0 ===
                    a) return c;
                if (isNaN(a)) return k.value;
                if (0 > a) return k.num;
                a = Math.floor(a);
                return a >= c.length ? REPT("0", a - c.length) + c : k.num
            };
            g.HEX2DEC = function(b) {
                if (!/^[0-9A-Fa-f]{1,10}$/.test(b)) return k.num;
                b = parseInt(b, 16);
                return 549755813888 <= b ? b - 1099511627776 : b
            };
            g.HEX2OCT = function(b, a) {
                if (!/^[0-9A-Fa-f]{1,10}$/.test(b)) return k.num;
                b = parseInt(b, 16);
                if (536870911 < b && 0xffe0000000 > b) return k.num;
                if (0xffe0000000 <= b) return (b - 0xffc0000000).toString(8);
                b = b.toString(8);
                if (void 0 === a) return b;
                if (isNaN(a)) return k.value;
                if (0 > a) return k.num;
                a = Math.floor(a);
                return a >= b.length ? REPT("0", a - b.length) + b : k.num
            };
            g.IMABS = function(b) {
                var a = g.IMREAL(b);
                b = g.IMAGINARY(b);
                return e.anyIsError(a, b) ? k.value : Math.sqrt(Math.pow(a, 2) + Math.pow(b, 2))
            };
            g.IMAGINARY = function(b) {
                if (void 0 === b || !0 === b || !1 === b) return k.value;
                if (0 === b || "0" === b) return 0;
                if (0 <= ["i", "j"].indexOf(b)) return 1;
                b = b.replace("+i", "+1i").replace("-i", "-1i").replace("+j", "+1j").replace("-j", "-1j");
                var a = b.indexOf("+"),
                    c = b.indexOf("-");
                0 === a && (a = b.indexOf("+", 1));
                0 === c &&
                    (c = b.indexOf("-", 1));
                var d = b.substring(b.length - 1, b.length);
                d = "i" === d || "j" === d;
                return 0 <= a || 0 <= c ? d ? 0 <= a ? isNaN(b.substring(0, a)) || isNaN(b.substring(a + 1, b.length - 1)) ? k.num : Number(b.substring(a + 1, b.length - 1)) : isNaN(b.substring(0, c)) || isNaN(b.substring(c + 1, b.length - 1)) ? k.num : -Number(b.substring(c + 1, b.length - 1)) : k.num : d ? isNaN(b.substring(0, b.length - 1)) ? k.num : b.substring(0, b.length - 1) : isNaN(b) ? k.num : 0
            };
            g.IMARGUMENT = function(b) {
                var a = g.IMREAL(b);
                b = g.IMAGINARY(b);
                return e.anyIsError(a, b) ? k.value : 0 === a && 0 ===
                    b ? k.div0 : 0 === a && 0 < b ? Math.PI / 2 : 0 === a && 0 > b ? -Math.PI / 2 : 0 === b && 0 < a ? 0 : 0 === b && 0 > a ? -Math.PI : 0 < a ? Math.atan(b / a) : 0 > a && 0 <= b ? Math.atan(b / a) + Math.PI : Math.atan(b / a) - Math.PI
            };
            g.IMCONJUGATE = function(b) {
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                if (e.anyIsError(a, c)) return k.value;
                var d = b.substring(b.length - 1);
                return 0 !== c ? g.COMPLEX(a, -c, "i" === d || "j" === d ? d : "i") : b
            };
            g.IMCOS = function(b) {
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                if (e.anyIsError(a, c)) return k.value;
                b = b.substring(b.length - 1);
                return g.COMPLEX(Math.cos(a) * (Math.exp(c) +
                    Math.exp(-c)) / 2, -Math.sin(a) * (Math.exp(c) - Math.exp(-c)) / 2, "i" === b || "j" === b ? b : "i")
            };
            g.IMCOSH = function(b) {
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                if (e.anyIsError(a, c)) return k.value;
                b = b.substring(b.length - 1);
                return g.COMPLEX(Math.cos(c) * (Math.exp(a) + Math.exp(-a)) / 2, Math.sin(c) * (Math.exp(a) - Math.exp(-a)) / 2, "i" === b || "j" === b ? b : "i")
            };
            g.IMCOT = function(b) {
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                return e.anyIsError(a, c) ? k.value : g.IMDIV(g.IMCOS(b), g.IMSIN(b))
            };
            g.IMDIV = function(b, a) {
                var c = g.IMREAL(b),
                    d = g.IMAGINARY(b),
                    f = g.IMREAL(a),
                    h = g.IMAGINARY(a);
                if (e.anyIsError(c, d, f, h)) return k.value;
                b = b.substring(b.length - 1);
                var l = a.substring(a.length - 1);
                a = "i";
                "j" === b ? a = "j" : "j" === l && (a = "j");
                if (0 === f && 0 === h) return k.num;
                b = f * f + h * h;
                return g.COMPLEX((c * f + d * h) / b, (d * f - c * h) / b, a)
            };
            g.IMEXP = function(b) {
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                if (e.anyIsError(a, c)) return k.value;
                b = b.substring(b.length - 1);
                a = Math.exp(a);
                return g.COMPLEX(a * Math.cos(c), a * Math.sin(c), "i" === b || "j" === b ? b : "i")
            };
            g.IMLN = function(b) {
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                if (e.anyIsError(a, c)) return k.value;
                b = b.substring(b.length - 1);
                return g.COMPLEX(Math.log(Math.sqrt(a * a + c * c)), Math.atan(c / a), "i" === b || "j" === b ? b : "i")
            };
            g.IMLOG10 = function(b) {
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                if (e.anyIsError(a, c)) return k.value;
                b = b.substring(b.length - 1);
                return g.COMPLEX(Math.log(Math.sqrt(a * a + c * c)) / Math.log(10), Math.atan(c / a) / Math.log(10), "i" === b || "j" === b ? b : "i")
            };
            g.IMLOG2 = function(b) {
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                if (e.anyIsError(a, c)) return k.value;
                b = b.substring(b.length - 1);
                return g.COMPLEX(Math.log(Math.sqrt(a *
                    a + c * c)) / Math.log(2), Math.atan(c / a) / Math.log(2), "i" === b || "j" === b ? b : "i")
            };
            g.IMPOWER = function(b, a) {
                a = e.parseNumber(a);
                var c = g.IMREAL(b),
                    d = g.IMAGINARY(b);
                if (e.anyIsError(a, c, d)) return k.value;
                c = b.substring(b.length - 1);
                c = "i" === c || "j" === c ? c : "i";
                d = Math.pow(g.IMABS(b), a);
                b = g.IMARGUMENT(b);
                return g.COMPLEX(d * Math.cos(a * b), d * Math.sin(a * b), c)
            };
            g.IMPRODUCT = function() {
                for (var b = arguments[0], a = 1; a < arguments.length; a++) {
                    var c = g.IMREAL(b);
                    b = g.IMAGINARY(b);
                    var d = g.IMREAL(arguments[a]),
                        f = g.IMAGINARY(arguments[a]);
                    if (e.anyIsError(c, b, d, f)) return k.value;
                    b = g.COMPLEX(c * d - b * f, c * f + b * d)
                }
                return b
            };
            g.IMREAL = function(b) {
                if (void 0 === b || !0 === b || !1 === b) return k.value;
                if (0 === b || "0" === b || 0 <= "i +i 1i +1i -i -1i j +j 1j +1j -j -1j".split(" ").indexOf(b)) return 0;
                var a = b.indexOf("+"),
                    c = b.indexOf("-");
                0 === a && (a = b.indexOf("+", 1));
                0 === c && (c = b.indexOf("-", 1));
                var d = b.substring(b.length - 1, b.length);
                d = "i" === d || "j" === d;
                return 0 <= a || 0 <= c ? d ? 0 <= a ? isNaN(b.substring(0, a)) || isNaN(b.substring(a + 1, b.length - 1)) ? k.num : Number(b.substring(0,
                    a)) : isNaN(b.substring(0, c)) || isNaN(b.substring(c + 1, b.length - 1)) ? k.num : Number(b.substring(0, c)) : k.num : d ? isNaN(b.substring(0, b.length - 1)) ? k.num : 0 : isNaN(b) ? k.num : b
            };
            g.IMSEC = function(b) {
                if (!0 === b || !1 === b) return k.value;
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                return e.anyIsError(a, c) ? k.value : g.IMDIV("1", g.IMCOS(b))
            };
            g.IMSECH = function(b) {
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                return e.anyIsError(a, c) ? k.value : g.IMDIV("1", g.IMCOSH(b))
            };
            g.IMSIN = function(b) {
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                if (e.anyIsError(a, c)) return k.value;
                b = b.substring(b.length - 1);
                return g.COMPLEX(Math.sin(a) * (Math.exp(c) + Math.exp(-c)) / 2, Math.cos(a) * (Math.exp(c) - Math.exp(-c)) / 2, "i" === b || "j" === b ? b : "i")
            };
            g.IMSINH = function(b) {
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                if (e.anyIsError(a, c)) return k.value;
                b = b.substring(b.length - 1);
                return g.COMPLEX(Math.cos(c) * (Math.exp(a) - Math.exp(-a)) / 2, Math.sin(c) * (Math.exp(a) + Math.exp(-a)) / 2, "i" === b || "j" === b ? b : "i")
            };
            g.IMSQRT = function(b) {
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                if (e.anyIsError(a, c)) return k.value;
                a = b.substring(b.length -
                    1);
                a = "i" === a || "j" === a ? a : "i";
                c = Math.sqrt(g.IMABS(b));
                b = g.IMARGUMENT(b);
                return g.COMPLEX(c * Math.cos(b / 2), c * Math.sin(b / 2), a)
            };
            g.IMCSC = function(b) {
                if (!0 === b || !1 === b) return k.value;
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                return e.anyIsError(a, c) ? k.num : g.IMDIV("1", g.IMSIN(b))
            };
            g.IMCSCH = function(b) {
                if (!0 === b || !1 === b) return k.value;
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                return e.anyIsError(a, c) ? k.num : g.IMDIV("1", g.IMSINH(b))
            };
            g.IMSUB = function(b, a) {
                var c = this.IMREAL(b),
                    d = this.IMAGINARY(b),
                    f = this.IMREAL(a),
                    h = this.IMAGINARY(a);
                if (e.anyIsError(c, d, f, h)) return k.value;
                b = b.substring(b.length - 1);
                a = a.substring(a.length - 1);
                var l = "i";
                "j" === b ? l = "j" : "j" === a && (l = "j");
                return this.COMPLEX(c - f, d - h, l)
            };
            g.IMSUM = function() {
                for (var b = e.flatten(arguments), a = b[0], c = 1; c < b.length; c++) {
                    var d = this.IMREAL(a);
                    a = this.IMAGINARY(a);
                    var f = this.IMREAL(b[c]),
                        h = this.IMAGINARY(b[c]);
                    if (e.anyIsError(d, a, f, h)) return k.value;
                    a = this.COMPLEX(d + f, a + h)
                }
                return a
            };
            g.IMTAN = function(b) {
                if (!0 === b || !1 === b) return k.value;
                var a = g.IMREAL(b),
                    c = g.IMAGINARY(b);
                return e.anyIsError(a,
                    c) ? k.value : this.IMDIV(this.IMSIN(b), this.IMCOS(b))
            };
            g.OCT2BIN = function(b, a) {
                if (!/^[0-7]{1,10}$/.test(b)) return k.num;
                var c = 10 === b.length && "7" === b.substring(0, 1) ? !0 : !1;
                b = c ? parseInt(b, 8) - 1073741824 : parseInt(b, 8);
                if (-512 > b || 511 < b) return k.num;
                if (c) return "1" + REPT("0", 9 - (512 + b).toString(2).length) + (512 + b).toString(2);
                c = b.toString(2);
                if ("undefined" === typeof a) return c;
                if (isNaN(a)) return k.value;
                if (0 > a) return k.num;
                a = Math.floor(a);
                return a >= c.length ? REPT("0", a - c.length) + c : k.num
            };
            g.OCT2DEC = function(b) {
                if (!/^[0-7]{1,10}$/.test(b)) return k.num;
                b = parseInt(b, 8);
                return 536870912 <= b ? b - 1073741824 : b
            };
            g.OCT2HEX = function(b, a) {
                if (!/^[0-7]{1,10}$/.test(b)) return k.num;
                b = parseInt(b, 8);
                if (536870912 <= b) return "ff" + (b + 3221225472).toString(16);
                b = b.toString(16);
                if (void 0 === a) return b;
                if (isNaN(a)) return k.value;
                if (0 > a) return k.num;
                a = Math.floor(a);
                return a >= b.length ? REPT("0", a - b.length) + b : k.num
            };
            return g
        }();
        y.financial = function() {
            function g(c) {
                return c && c.getTime && !isNaN(c.getTime())
            }

            function b(c) {
                return c instanceof Date ? c : new Date(c)
            }
            var a = {
                ACCRINT: function(c,
                    d, f, h, l, m, n) {
                    c = b(c);
                    d = b(d);
                    f = b(f);
                    return g(c) && g(d) && g(f) ? 0 >= h || 0 >= l || -1 === [1, 2, 4].indexOf(m) || -1 === [0, 1, 2, 3, 4].indexOf(n) || f <= c ? "#NUM!" : (l || 0) * h * YEARFRAC(c, f, n || 0) : "#VALUE!"
                },
                ACCRINTM: null,
                AMORDEGRC: null,
                AMORLINC: null,
                COUPDAYBS: null,
                COUPDAYS: null,
                COUPDAYSNC: null,
                COUPNCD: null,
                COUPNUM: null,
                COUPPCD: null,
                CUMIPMT: function(c, d, f, h, l, m) {
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    if (e.anyIsError(c, d, f)) return k.value;
                    if (0 >= c || 0 >= d || 0 >= f || 1 > h || 1 > l || h > l || 0 !== m && 1 !== m) return k.num;
                    d = a.PMT(c, d,
                        f, 0, m);
                    var n = 0;
                    1 === h && 0 === m && (n = -f, h++);
                    for (; h <= l; h++) n = 1 === m ? n + (a.FV(c, h - 2, d, f, 1) - d) : n + a.FV(c, h - 1, d, f, 0);
                    return n * c
                },
                CUMPRINC: function(c, d, f, h, l, m) {
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    if (e.anyIsError(c, d, f)) return k.value;
                    if (0 >= c || 0 >= d || 0 >= f || 1 > h || 1 > l || h > l || 0 !== m && 1 !== m) return k.num;
                    d = a.PMT(c, d, f, 0, m);
                    var n = 0;
                    1 === h && (n = 0 === m ? d + f * c : d, h++);
                    for (; h <= l; h++) n = 0 < m ? n + (d - (a.FV(c, h - 2, d, f, 1) - d) * c) : n + (d - a.FV(c, h - 1, d, f, 0) * c);
                    return n
                },
                DB: function(c, d, f, h, l) {
                    l = void 0 === l ? 12 : l;
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    h = e.parseNumber(h);
                    l = e.parseNumber(l);
                    if (e.anyIsError(c, d, f, h, l)) return k.value;
                    if (0 > c || 0 > d || 0 > f || 0 > h || -1 === [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].indexOf(l) || h > f) return k.num;
                    if (d >= c) return 0;
                    d = (1 - Math.pow(d / c, 1 / f)).toFixed(3);
                    for (var m = l = c * d * l / 12, n = 0, p = h === f ? f - 1 : h, q = 2; q <= p; q++) n = (c - m) * d, m += n;
                    return 1 === h ? l : h === f ? (c - m) * d : n
                },
                DDB: function(c, d, f, h, l) {
                    l = void 0 === l ? 2 : l;
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    h = e.parseNumber(h);
                    l = e.parseNumber(l);
                    if (e.anyIsError(c,
                            d, f, h, l)) return k.value;
                    if (0 > c || 0 > d || 0 > f || 0 > h || 0 >= l || h > f) return k.num;
                    if (d >= c) return 0;
                    for (var m = 0, n = 0, p = 1; p <= h; p++) n = Math.min(l / f * (c - m), c - d - m), m += n;
                    return n
                },
                DISC: null,
                DOLLARDE: function(c, d) {
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    if (e.anyIsError(c, d)) return k.value;
                    if (0 > d) return k.num;
                    if (0 <= d && 1 > d) return k.div0;
                    d = parseInt(d, 10);
                    var f = parseInt(c, 10);
                    f += c % 1 * Math.pow(10, Math.ceil(Math.log(d) / Math.LN10)) / d;
                    c = Math.pow(10, Math.ceil(Math.log(d) / Math.LN2) + 1);
                    return f = Math.round(f * c) / c
                },
                DOLLARFR: function(c,
                    d) {
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    if (e.anyIsError(c, d)) return k.value;
                    if (0 > d) return k.num;
                    if (0 <= d && 1 > d) return k.div0;
                    d = parseInt(d, 10);
                    var f = parseInt(c, 10);
                    return f += c % 1 * Math.pow(10, -Math.ceil(Math.log(d) / Math.LN10)) * d
                },
                DURATION: null,
                EFFECT: function(c, d) {
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    if (e.anyIsError(c, d)) return k.value;
                    if (0 >= c || 1 > d) return k.num;
                    d = parseInt(d, 10);
                    return Math.pow(1 + c / d, d) - 1
                },
                FV: function(c, d, f, h, l) {
                    h = h || 0;
                    l = l || 0;
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    h =
                        e.parseNumber(h);
                    l = e.parseNumber(l);
                    if (e.anyIsError(c, d, f, h, l)) return k.value;
                    0 === c ? c = h + f * d : (d = Math.pow(1 + c, d), c = 1 === l ? h * d + f * (1 + c) * (d - 1) / c : h * d + f * (d - 1) / c);
                    return -c
                },
                FVSCHEDULE: function(c, d) {
                    c = e.parseNumber(c);
                    d = e.parseNumberArray(e.flatten(d));
                    if (e.anyIsError(c, d)) return k.value;
                    for (var f = d.length, h = 0; h < f; h++) c *= 1 + d[h];
                    return c
                },
                INTRATE: null,
                IPMT: function(c, d, f, h, l, m) {
                    l = l || 0;
                    m = m || 0;
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    h = e.parseNumber(h);
                    l = e.parseNumber(l);
                    m = e.parseNumber(m);
                    if (e.anyIsError(c,
                            d, f, h, l, m)) return k.value;
                    f = a.PMT(c, f, h, l, m);
                    return (1 === d ? 1 === m ? 0 : -h : 1 === m ? a.FV(c, d - 2, f, h, 1) - f : a.FV(c, d - 1, f, h, 0)) * c
                },
                IRR: function(c, d) {
                    d = d || 0;
                    c = e.parseNumberArray(e.flatten(c));
                    d = e.parseNumber(d);
                    if (e.anyIsError(c, d)) return k.value;
                    for (var f = function(q, t, u) {
                            u += 1;
                            for (var r = q[0], v = 1; v < q.length; v++) r += q[v] / Math.pow(u, (t[v] - t[0]) / 365);
                            return r
                        }, h = function(q, t, u) {
                            u += 1;
                            for (var r = 0, v = 1; v < q.length; v++) {
                                var E = (t[v] - t[0]) / 365;
                                r -= E * q[v] / Math.pow(u, E + 1)
                            }
                            return r
                        }, l = [], m = !1, n = !1, p = 0; p < c.length; p++) l[p] = 0 === p ?
                        0 : l[p - 1] + 365, 0 < c[p] && (m = !0), 0 > c[p] && (n = !0);
                    if (!m || !n) return k.num;
                    d = void 0 === d ? .1 : d;
                    m = !0;
                    do p = f(c, l, d), m = d - p / h(c, l, d), n = Math.abs(m - d), d = m, m = 1E-10 < n && 1E-10 < Math.abs(p); while (m);
                    return d
                },
                ISPMT: function(c, d, f, h) {
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    h = e.parseNumber(h);
                    return e.anyIsError(c, d, f, h) ? k.value : h * c * (d / f - 1)
                },
                MDURATION: null,
                MIRR: function(c, d, f) {
                    c = e.parseNumberArray(e.flatten(c));
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    if (e.anyIsError(c, d, f)) return k.value;
                    for (var h = c.length,
                            l = [], m = [], n = 0; n < h; n++) 0 > c[n] ? l.push(c[n]) : m.push(c[n]);
                    c = -a.NPV(f, m) * Math.pow(1 + f, h - 1);
                    d = a.NPV(d, l) * (1 + d);
                    return Math.pow(c / d, 1 / (h - 1)) - 1
                },
                NOMINAL: function(c, d) {
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    if (e.anyIsError(c, d)) return k.value;
                    if (0 >= c || 1 > d) return k.num;
                    d = parseInt(d, 10);
                    return (Math.pow(c + 1, 1 / d) - 1) * d
                },
                NPER: function(c, d, f, h, l) {
                    l = void 0 === l ? 0 : l;
                    h = void 0 === h ? 0 : h;
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    h = e.parseNumber(h);
                    l = e.parseNumber(l);
                    return e.anyIsError(c, d, f, h, l) ? k.value :
                        Math.log((d * (1 + c * l) - h * c) / (f * c + d * (1 + c * l))) / Math.log(1 + c)
                },
                NPV: function() {
                    var c = e.parseNumberArray(e.flatten(arguments));
                    if (c instanceof Error) return c;
                    for (var d = c[0], f = 0, h = 1; h < c.length; h++) f += c[h] / Math.pow(1 + d, h);
                    return f
                },
                ODDFPRICE: null,
                ODDFYIELD: null,
                ODDLPRICE: null,
                ODDLYIELD: null,
                PDURATION: function(c, d, f) {
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    return e.anyIsError(c, d, f) ? k.value : 0 >= c ? k.num : (Math.log(f) - Math.log(d)) / Math.log(1 + c)
                },
                PMT: function(c, d, f, h, l) {
                    h = h || 0;
                    l = l || 0;
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    h = e.parseNumber(h);
                    l = e.parseNumber(l);
                    if (e.anyIsError(c, d, f, h, l)) return k.value;
                    0 === c ? c = (f + h) / d : (d = Math.pow(1 + c, d), c = 1 === l ? (h * c / (d - 1) + f * c / (1 - 1 / d)) / (1 + c) : h * c / (d - 1) + f * c / (1 - 1 / d));
                    return -c
                },
                PPMT: function(c, d, f, h, l, m) {
                    l = l || 0;
                    m = m || 0;
                    c = e.parseNumber(c);
                    f = e.parseNumber(f);
                    h = e.parseNumber(h);
                    l = e.parseNumber(l);
                    m = e.parseNumber(m);
                    return e.anyIsError(c, f, h, l, m) ? k.value : a.PMT(c, f, h, l, m) - a.IPMT(c, d, f, h, l, m)
                },
                PRICE: null,
                PRICEDISC: null,
                PRICEMAT: null,
                PV: function(c, d, f, h, l) {
                    h =
                        h || 0;
                    l = l || 0;
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    h = e.parseNumber(h);
                    l = e.parseNumber(l);
                    return e.anyIsError(c, d, f, h, l) ? k.value : 0 === c ? -f * d - h : ((1 - Math.pow(1 + c, d)) / c * f * (1 + c * l) - h) / Math.pow(1 + c, d)
                },
                RATE: function(c, d, f, h, l, m) {
                    m = void 0 === m ? .01 : m;
                    h = void 0 === h ? 0 : h;
                    l = void 0 === l ? 0 : l;
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    h = e.parseNumber(h);
                    l = e.parseNumber(l);
                    m = e.parseNumber(m);
                    if (e.anyIsError(c, d, f, h, l, m)) return k.value;
                    for (var n = 0, p = !1; 100 > n && !p;) {
                        var q = Math.pow(m + 1, c),
                            t = Math.pow(m + 1, c - 1);
                        q = m - (h + q * f + d * (q - 1) * (m * l + 1) / m) / (c * t * f - d * (q - 1) * (m * l + 1) / Math.pow(m, 2) + (c * d * t * (m * l + 1) / m + d * (q - 1) * l / m));
                        1E-6 > Math.abs(q - m) && (p = !0);
                        n++;
                        m = q
                    }
                    return p ? m : Number.NaN + m
                },
                RECEIVED: null,
                RRI: function(c, d, f) {
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    return e.anyIsError(c, d, f) ? k.value : 0 === c || 0 === d ? k.num : Math.pow(f / d, 1 / c) - 1
                },
                SLN: function(c, d, f) {
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    return e.anyIsError(c, d, f) ? k.value : 0 === f ? k.num : (c - d) / f
                },
                SYD: function(c, d, f, h) {
                    c =
                        e.parseNumber(c);
                    d = e.parseNumber(d);
                    f = e.parseNumber(f);
                    h = e.parseNumber(h);
                    if (e.anyIsError(c, d, f, h)) return k.value;
                    if (0 === f || 1 > h || h > f) return k.num;
                    h = parseInt(h, 10);
                    return (c - d) * (f - h + 1) * 2 / (f * (f + 1))
                },
                TBILLEQ: function(c, d, f) {
                    c = e.parseDate(c);
                    d = e.parseDate(d);
                    f = e.parseNumber(f);
                    return e.anyIsError(c, d, f) ? k.value : 0 >= f || c > d || 31536E6 < d - c ? k.num : 365 * f / (360 - f * DAYS360(c, d, !1))
                },
                TBILLPRICE: function(c, d, f) {
                    c = e.parseDate(c);
                    d = e.parseDate(d);
                    f = e.parseNumber(f);
                    return e.anyIsError(c, d, f) ? k.value : 0 >= f || c > d || 31536E6 <
                        d - c ? k.num : 100 * (1 - f * DAYS360(c, d, !1) / 360)
                },
                TBILLYIELD: function(c, d, f) {
                    c = e.parseDate(c);
                    d = e.parseDate(d);
                    f = e.parseNumber(f);
                    return e.anyIsError(c, d, f) ? k.value : 0 >= f || c > d || 31536E6 < d - c ? k.num : 360 * (100 - f) / (f * DAYS360(c, d, !1))
                },
                VDB: null,
                XIRR: function(c, d, f) {
                    c = e.parseNumberArray(e.flatten(c));
                    d = e.parseDateArray(e.flatten(d));
                    f = e.parseNumber(f);
                    if (e.anyIsError(c, d, f)) return k.value;
                    for (var h = function(q, t, u) {
                            u += 1;
                            for (var r = q[0], v = 1; v < q.length; v++) r += q[v] / Math.pow(u, DAYS(t[v], t[0]) / 365);
                            return r
                        }, l = function(q,
                            t, u) {
                            u += 1;
                            for (var r = 0, v = 1; v < q.length; v++) {
                                var E = DAYS(t[v], t[0]) / 365;
                                r -= E * q[v] / Math.pow(u, E + 1)
                            }
                            return r
                        }, m = !1, n = !1, p = 0; p < c.length; p++) 0 < c[p] && (m = !0), 0 > c[p] && (n = !0);
                    if (!m || !n) return k.num;
                    f = f || .1;
                    m = !0;
                    do p = h(c, d, f), m = f - p / l(c, d, f), n = Math.abs(m - f), f = m, m = 1E-10 < n && 1E-10 < Math.abs(p); while (m);
                    return f
                },
                XNPV: function(c, d, f) {
                    c = e.parseNumber(c);
                    d = e.parseNumberArray(e.flatten(d));
                    f = e.parseDateArray(e.flatten(f));
                    if (e.anyIsError(c, d, f)) return k.value;
                    for (var h = 0, l = 0; l < d.length; l++) h += d[l] / Math.pow(1 + c, DAYS(f[l],
                        f[0]) / 365);
                    return h
                },
                YIELD: null,
                YIELDDISC: null,
                YIELDMAT: null
            };
            return a
        }();
        y.information = function() {
            var g = {
                CELL: null,
                ERROR: {}
            };
            g.ERROR.TYPE = function(b) {
                switch (b) {
                    case k.nil:
                        return 1;
                    case k.div0:
                        return 2;
                    case k.value:
                        return 3;
                    case k.ref:
                        return 4;
                    case k.name:
                        return 5;
                    case k.num:
                        return 6;
                    case k.na:
                        return 7;
                    case k.data:
                        return 8
                }
                return k.na
            };
            g.INFO = null;
            g.ISBLANK = function(b) {
                return null === b || void 0 === b || "" === b
            };
            g.ISBINARY = function(b) {
                return /^[01]{1,10}$/.test(b)
            };
            g.ISERR = function(b) {
                return 0 <= [k.value,
                    k.ref, k.div0, k.num, k.name, k.nil
                ].indexOf(b) || "number" === typeof b && (isNaN(b) || !isFinite(b))
            };
            g.ISERROR = function(b) {
                return g.ISERR(b) || b === k.na
            };
            g.ISEVEN = function(b) {
                return Math.floor(Math.abs(b)) & 1 ? !1 : !0
            };
            g.ISFORMULA = null;
            g.ISLOGICAL = function(b) {
                return !0 === b || !1 === b
            };
            g.ISNA = function(b) {
                return b === k.na
            };
            g.ISNONTEXT = function(b) {
                return "string" !== typeof b
            };
            g.ISNUMBER = function(b) {
                return "number" === typeof b && !isNaN(b) && isFinite(b)
            };
            g.ISODD = function(b) {
                return Math.floor(Math.abs(b)) & 1 ? !0 : !1
            };
            g.ISREF = null;
            g.ISTEXT = function(b) {
                return "string" === typeof b
            };
            g.N = function(b) {
                return this.ISNUMBER(b) ? b : b instanceof Date ? b.getTime() : !0 === b ? 1 : !1 === b ? 0 : this.ISERROR(b) ? b : 0
            };
            g.NA = function() {
                return k.na
            };
            g.SHEET = null;
            g.SHEETS = null;
            g.TYPE = function(b) {
                if (this.ISNUMBER(b)) return 1;
                if (this.ISTEXT(b)) return 2;
                if (this.ISLOGICAL(b)) return 4;
                if (this.ISERROR(b)) return 16;
                if (Array.isArray(b)) return 64
            };
            return g
        }();
        y.logical = function() {
            return {
                AND: function() {
                    for (var g = e.flatten(arguments), b = !0, a = 0; a < g.length; a++) g[a] || (b = !1);
                    return b
                },
                CHOOSE: function() {
                    if (2 > arguments.length) return k.na;
                    var g = arguments[0];
                    return 1 > g || 254 < g || arguments.length < g + 1 ? k.value : arguments[g]
                },
                FALSE: function() {
                    return !1
                },
                IF: function(g, b, a) {
                    return g ? b : a
                },
                IFERROR: function(g, b) {
                    return ISERROR(g) ? b : g
                },
                IFNA: function(g, b) {
                    return g === k.na ? b : g
                },
                NOT: function(g) {
                    return !g
                },
                OR: function() {
                    for (var g = e.flatten(arguments), b = !1, a = 0; a < g.length; a++) g[a] && (b = !0);
                    return b
                },
                TRUE: function() {
                    return !0
                },
                XOR: function() {
                    for (var g = e.flatten(arguments), b = 0, a = 0; a < g.length; a++) g[a] &&
                        b++;
                    return Math.floor(Math.abs(b)) & 1 ? !0 : !1
                },
                SWITCH: function() {
                    if (0 < arguments.length) {
                        var g = arguments[0],
                            b = arguments.length - 1,
                            a = Math.floor(b / 2),
                            c = !1;
                        b = 0 === b % 2 ? null : arguments[arguments.length - 1];
                        if (a)
                            for (var d = 0; d < a; d++)
                                if (g === arguments[2 * d + 1]) {
                                    var f = arguments[2 * d + 2];
                                    c = !0;
                                    break
                                }! c && b && (f = b)
                    }
                    return f
                }
            }
        }();
        y.math = function() {
            var g = {
                ABS: function(a) {
                    a = e.parseNumber(a);
                    return a instanceof Error ? a : Math.abs(e.parseNumber(a))
                },
                ACOS: function(a) {
                    a = e.parseNumber(a);
                    return a instanceof Error ? a : Math.acos(a)
                },
                ACOSH: function(a) {
                    a =
                        e.parseNumber(a);
                    return a instanceof Error ? a : Math.log(a + Math.sqrt(a * a - 1))
                },
                ACOT: function(a) {
                    a = e.parseNumber(a);
                    return a instanceof Error ? a : Math.atan(1 / a)
                },
                ACOTH: function(a) {
                    a = e.parseNumber(a);
                    return a instanceof Error ? a : .5 * Math.log((a + 1) / (a - 1))
                },
                AGGREGATE: null,
                ARABIC: function(a) {
                    if (!/^M*(?:D?C{0,3}|C[MD])(?:L?X{0,3}|X[CL])(?:V?I{0,3}|I[XV])$/.test(a)) return k.value;
                    var c = 0;
                    a.replace(/[MDLV]|C[MD]?|X[CL]?|I[XV]?/g, function(d) {
                        c += {
                            M: 1E3,
                            CM: 900,
                            D: 500,
                            CD: 400,
                            C: 100,
                            XC: 90,
                            L: 50,
                            XL: 40,
                            X: 10,
                            IX: 9,
                            V: 5,
                            IV: 4,
                            I: 1
                        } [d]
                    });
                    return c
                },
                ASIN: function(a) {
                    a = e.parseNumber(a);
                    return a instanceof Error ? a : Math.asin(a)
                },
                ASINH: function(a) {
                    a = e.parseNumber(a);
                    return a instanceof Error ? a : Math.log(a + Math.sqrt(a * a + 1))
                },
                ATAN: function(a) {
                    a = e.parseNumber(a);
                    return a instanceof Error ? a : Math.atan(a)
                },
                ATAN2: function(a, c) {
                    a = e.parseNumber(a);
                    c = e.parseNumber(c);
                    return e.anyIsError(a, c) ? k.value : Math.atan2(a, c)
                },
                ATANH: function(a) {
                    a = e.parseNumber(a);
                    return a instanceof Error ? a : Math.log((1 + a) / (1 - a)) / 2
                },
                BASE: function(a, c, d) {
                    d = d || 0;
                    a = e.parseNumber(a);
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    if (e.anyIsError(a, c, d)) return k.value;
                    d = void 0 === d ? 0 : d;
                    a = a.toString(c);
                    return Array(Math.max(d + 1 - a.length, 0)).join("0") + a
                },
                CEILING: function(a, c, d) {
                    c = void 0 === c ? 1 : c;
                    d = void 0 === d ? 0 : d;
                    a = e.parseNumber(a);
                    c = e.parseNumber(c);
                    d = e.parseNumber(d);
                    if (e.anyIsError(a, c, d)) return k.value;
                    if (0 === c) return 0;
                    c = Math.abs(c);
                    return 0 <= a ? Math.ceil(a / c) * c : 0 === d ? -1 * Math.floor(Math.abs(a) / c) * c : -1 * Math.ceil(Math.abs(a) / c) * c
                }
            };
            g.CEILING.MATH = g.CEILING;
            g.CEILING.PRECISE = g.CEILING;
            g.COMBIN =
                function(a, c) {
                    a = e.parseNumber(a);
                    c = e.parseNumber(c);
                    return e.anyIsError(a, c) ? k.value : g.FACT(a) / (g.FACT(c) * g.FACT(a - c))
                };
            g.COMBINA = function(a, c) {
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(a, c) ? k.value : 0 === a && 0 === c ? 1 : g.COMBIN(a + c - 1, a - 1)
            };
            g.COS = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : Math.cos(a)
            };
            g.COSH = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : (Math.exp(a) + Math.exp(-a)) / 2
            };
            g.COT = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : 1 / Math.tan(a)
            };
            g.COTH = function(a) {
                a = e.parseNumber(a);
                if (a instanceof Error) return a;
                a = Math.exp(2 * a);
                return (a + 1) / (a - 1)
            };
            g.CSC = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : 1 / Math.sin(a)
            };
            g.CSCH = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : 2 / (Math.exp(a) - Math.exp(-a))
            };
            g.DECIMAL = function(a, c) {
                return 1 > arguments.length ? k.value : parseInt(a, c)
            };
            g.DEGREES = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : 180 * a / Math.PI
            };
            g.EVEN = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ?
                    a : g.CEILING(a, -2, -1)
            };
            g.EXP = Math.exp;
            var b = [];
            g.FACT = function(a) {
                a = e.parseNumber(a);
                if (a instanceof Error) return a;
                a = Math.floor(a);
                if (0 === a || 1 === a) return 1;
                0 < b[a] || (b[a] = g.FACT(a - 1) * a);
                return b[a]
            };
            g.FACTDOUBLE = function(a) {
                a = e.parseNumber(a);
                if (a instanceof Error) return a;
                a = Math.floor(a);
                return 0 >= a ? 1 : a * g.FACTDOUBLE(a - 2)
            };
            g.FLOOR = function(a, c, d) {
                c = void 0 === c ? 1 : c;
                d = void 0 === d ? 0 : d;
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                d = e.parseNumber(d);
                if (e.anyIsError(a, c, d)) return k.value;
                if (0 === c) return 0;
                c = Math.abs(c);
                return 0 <= a ? Math.floor(a / c) * c : 0 === d ? -1 * Math.ceil(Math.abs(a) / c) * c : -1 * Math.floor(Math.abs(a) / c) * c
            };
            g.FLOOR.MATH = g.FLOOR;
            g.GCD = null;
            g.INT = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : Math.floor(a)
            };
            g.LCM = function() {
                var a = e.parseNumberArray(e.flatten(arguments));
                if (a instanceof Error) return a;
                for (var c, d, f, h = 1; void 0 !== (f = a.pop());)
                    for (; 1 < f;) {
                        if (f % 2) {
                            c = 3;
                            for (d = Math.floor(Math.sqrt(f)); c <= d && f % c; c += 2);
                            d = c <= d ? c : f
                        } else d = 2;
                        f /= d;
                        h *= d;
                        for (c = a.length; c; 0 === a[--c] % d && 1 === (a[c] /= d) && a.splice(c,
                                1));
                    }
                return h
            };
            g.LN = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : Math.log(a)
            };
            g.LOG = function(a, c) {
                a = e.parseNumber(a);
                c = void 0 === c ? 10 : e.parseNumber(c);
                return e.anyIsError(a, c) ? k.value : Math.log(a) / Math.log(c)
            };
            g.LOG10 = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : 0 === a ? k.num : Math.log(a) / Math.log(10)
            };
            g.MDETERM = null;
            g.MINVERSE = null;
            g.MMULT = null;
            g.MOD = function(a, c) {
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                if (e.anyIsError(a, c)) return k.value;
                if (0 === c) return k.div0;
                a = Math.abs(a %
                    c);
                return 0 < c ? a : -a
            };
            g.MROUND = function(a, c) {
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(a, c) ? k.value : 0 > a * c ? k.num : Math.round(a / c) * c
            };
            g.MULTINOMIAL = function() {
                var a = e.parseNumberArray(e.flatten(arguments));
                if (a instanceof Error) return a;
                for (var c = 0, d = 1, f = 0; f < a.length; f++) c += a[f], d *= g.FACT(a[f]);
                return g.FACT(c) / d
            };
            g.MUNIT = null;
            g.ODD = function(a) {
                a = e.parseNumber(a);
                if (a instanceof Error) return a;
                var c = Math.ceil(Math.abs(a));
                c = c & 1 ? c : c + 1;
                return 0 < a ? c : -c
            };
            g.PI = function() {
                return Math.PI
            };
            g.POWER =
                function(a, c) {
                    a = e.parseNumber(a);
                    c = e.parseNumber(c);
                    if (e.anyIsError(a, c)) return k.value;
                    a = Math.pow(a, c);
                    return isNaN(a) ? k.num : a
                };
            g.PRODUCT = function() {
                var a = e.parseNumberArray(e.flatten(arguments));
                if (a instanceof Error) return a;
                for (var c = 1, d = 0; d < a.length; d++) c *= a[d];
                return c
            };
            g.QUOTIENT = function(a, c) {
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(a, c) ? k.value : parseInt(a / c, 10)
            };
            g.RADIANS = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : a * Math.PI / 180
            };
            g.RAND = function() {
                return Math.random()
            };
            g.RANDBETWEEN = function(a, c) {
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(a, c) ? k.value : a + Math.ceil((c - a + 1) * Math.random()) - 1
            };
            g.ROMAN = null;
            g.ROUND = function(a, c) {
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(a, c) ? k.value : Math.round(a * Math.pow(10, c)) / Math.pow(10, c)
            };
            g.ROUNDDOWN = function(a, c) {
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(a, c) ? k.value : (0 < a ? 1 : -1) * Math.floor(Math.abs(a) * Math.pow(10, c)) / Math.pow(10, c)
            };
            g.ROUNDUP = function(a, c) {
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(a, c) ? k.value : (0 < a ? 1 : -1) * Math.ceil(Math.abs(a) * Math.pow(10, c)) / Math.pow(10, c)
            };
            g.SEC = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : 1 / Math.cos(a)
            };
            g.SECH = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : 2 / (Math.exp(a) + Math.exp(-a))
            };
            g.SERIESSUM = function(a, c, d, f) {
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                d = e.parseNumber(d);
                f = e.parseNumberArray(f);
                if (e.anyIsError(a, c, d, f)) return k.value;
                for (var h = f[0] * Math.pow(a, c), l = 1; l < f.length; l++) h += f[l] * Math.pow(a, c + l * d);
                return h
            };
            g.SIGN = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : 0 > a ? -1 : 0 === a ? 0 : 1
            };
            g.SIN = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : Math.sin(a)
            };
            g.SINH = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : (Math.exp(a) - Math.exp(-a)) / 2
            };
            g.SQRT = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : 0 > a ? k.num : Math.sqrt(a)
            };
            g.SQRTPI = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : Math.sqrt(a * Math.PI)
            };
            g.SUBTOTAL = null;
            g.ADD = function(a, c) {
                if (2 !== arguments.length) return k.na;
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(a, c) ? k.value : a + c
            };
            g.MINUS = function(a, c) {
                if (2 !== arguments.length) return k.na;
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(a, c) ? k.value : a - c
            };
            g.DIVIDE = function(a, c) {
                if (2 !== arguments.length) return k.na;
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(a, c) ? k.value : 0 === c ? k.div0 : a / c
            };
            g.MULTIPLY = function(a, c) {
                if (2 !== arguments.length) return k.na;
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(a, c) ? k.value : a * c
            };
            g.GTE =
                function(a, c) {
                    if (2 !== arguments.length) return k.na;
                    a = e.parseNumber(a);
                    c = e.parseNumber(c);
                    return e.anyIsError(a, c) ? k.error : a >= c
                };
            g.LT = function(a, c) {
                if (2 !== arguments.length) return k.na;
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(a, c) ? k.error : a < c
            };
            g.LTE = function(a, c) {
                if (2 !== arguments.length) return k.na;
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(a, c) ? k.error : a <= c
            };
            g.EQ = function(a, c) {
                return 2 !== arguments.length ? k.na : a === c
            };
            g.NE = function(a, c) {
                return 2 !== arguments.length ? k.na :
                    a !== c
            };
            g.POW = function(a, c) {
                if (2 !== arguments.length) return k.na;
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(a, c) ? k.error : g.POWER(a, c)
            };
            g.SUM = function() {
                for (var a = 0, c = Object.keys(arguments), d = 0; d < c.length; ++d) {
                    var f = arguments[c[d]];
                    "number" === typeof f ? a += f : "string" === typeof f ? (f = parseFloat(f), !isNaN(f) && (a += f)) : Array.isArray(f) && (a += g.SUM.apply(null, f))
                }
                return a
            };
            g.SUMIF = function() {
                var a = e.argsToArray(arguments),
                    c = a.pop();
                a = e.parseNumberArray(e.flatten(a));
                if (a instanceof Error) return a;
                for (var d = 0, f = 0; f < a.length; f++) d += eval(a[f] + c) ? a[f] : 0;
                return d
            };
            g.SUMIFS = function() {
                var a = e.argsToArray(arguments),
                    c = e.parseNumberArray(e.flatten(a.shift()));
                if (c instanceof Error) return c;
                for (var d = c.length, f = a.length, h = 0, l = 0; l < d; l++) {
                    for (var m = c[l], n = "", p = 0; p < f; p += 2) n = isNaN(a[p][l]) ? n + ('"' + a[p][l] + '"' + a[p + 1]) : n + (a[p][l] + a[p + 1]), p !== f - 1 && (n += " && ");
                    n = n.slice(0, -4);
                    eval(n) && (h += m)
                }
                return h
            };
            g.SUMPRODUCT = null;
            g.SUMSQ = function() {
                var a = e.parseNumberArray(e.flatten(arguments));
                if (a instanceof Error) return a;
                for (var c = 0, d = a.length, f = 0; f < d; f++) c += ISNUMBER(a[f]) ? a[f] * a[f] : 0;
                return c
            };
            g.SUMX2MY2 = function(a, c) {
                a = e.parseNumberArray(e.flatten(a));
                c = e.parseNumberArray(e.flatten(c));
                if (e.anyIsError(a, c)) return k.value;
                for (var d = 0, f = 0; f < a.length; f++) d += a[f] * a[f] - c[f] * c[f];
                return d
            };
            g.SUMX2PY2 = function(a, c) {
                a = e.parseNumberArray(e.flatten(a));
                c = e.parseNumberArray(e.flatten(c));
                if (e.anyIsError(a, c)) return k.value;
                var d = 0;
                a = e.parseNumberArray(e.flatten(a));
                c = e.parseNumberArray(e.flatten(c));
                for (var f = 0; f < a.length; f++) d +=
                    a[f] * a[f] + c[f] * c[f];
                return d
            };
            g.SUMXMY2 = function(a, c) {
                a = e.parseNumberArray(e.flatten(a));
                c = e.parseNumberArray(e.flatten(c));
                if (e.anyIsError(a, c)) return k.value;
                var d = 0;
                a = e.flatten(a);
                c = e.flatten(c);
                for (var f = 0; f < a.length; f++) d += Math.pow(a[f] - c[f], 2);
                return d
            };
            g.TAN = function(a) {
                a = e.parseNumber(a);
                return a instanceof Error ? a : Math.tan(a)
            };
            g.TANH = function(a) {
                a = e.parseNumber(a);
                if (a instanceof Error) return a;
                a = Math.exp(2 * a);
                return (a - 1) / (a + 1)
            };
            g.TRUNC = function(a, c) {
                c = void 0 === c ? 0 : c;
                a = e.parseNumber(a);
                c =
                    e.parseNumber(c);
                return e.anyIsError(a, c) ? k.value : (0 < a ? 1 : -1) * Math.floor(Math.abs(a) * Math.pow(10, c)) / Math.pow(10, c)
            };
            return g
        }();
        y.misc = function() {
            var g = {
                UNIQUE: function() {
                    for (var b = [], a = 0; a < arguments.length; ++a) {
                        for (var c = !1, d = arguments[a], f = 0; f < b.length && !(c = b[f] === d); ++f);
                        c || b.push(d)
                    }
                    return b
                }
            };
            g.FLATTEN = e.flatten;
            g.ARGS2ARRAY = function() {
                return Array.prototype.slice.call(arguments, 0)
            };
            g.REFERENCE = function(b, a) {
                try {
                    var c = a.split(".");
                    for (a = 0; a < c.length; ++a) {
                        var d = c[a];
                        if ("]" === d[d.length - 1]) {
                            var f =
                                d.indexOf("["),
                                h = d.substring(f + 1, d.length - 1);
                            b = b[d.substring(0, f)][h]
                        } else b = b[d]
                    }
                    return b
                } catch (l) {}
            };
            g.JOIN = function(b, a) {
                return b.join(a)
            };
            g.NUMBERS = function() {
                return e.flatten(arguments).filter(function(b) {
                    return "number" === typeof b
                })
            };
            g.NUMERAL = null;
            return g
        }();
        y.text = function() {
            var g = {
                ASC: null,
                BAHTTEXT: null,
                CHAR: function(b) {
                    b = e.parseNumber(b);
                    return b instanceof Error ? b : String.fromCharCode(b)
                },
                CLEAN: function(b) {
                    return (b || "").replace(/[\0-\x1F]/g, "")
                },
                CODE: function(b) {
                    return (b || "").charCodeAt(0)
                },
                CONCATENATE: function() {
                    for (var b = e.flatten(arguments), a; - 1 < (a = b.indexOf(!0));) b[a] = "TRUE";
                    for (; - 1 < (a = b.indexOf(!1));) b[a] = "FALSE";
                    return b.join("")
                },
                DBCS: null,
                DOLLAR: null,
                EXACT: function(b, a) {
                    return b === a
                },
                FIND: function(b, a, c) {
                    return a ? a.indexOf(b, (void 0 === c ? 0 : c) - 1) + 1 : null
                },
                FIXED: null,
                HTML2TEXT: function(b) {
                    var a = "";
                    b && (b instanceof Array ? b.forEach(function(c) {
                        "" !== a && (a += "\n");
                        a += c.replace(/<(?:.|\n)*?>/gm, "")
                    }) : a = b.replace(/<(?:.|\n)*?>/gm, ""));
                    return a
                },
                LEFT: function(b, a) {
                    a = e.parseNumber(void 0 ===
                        a ? 1 : a);
                    return a instanceof Error || "string" !== typeof b ? k.value : b ? b.substring(0, a) : null
                },
                LEN: function(b) {
                    return 0 === arguments.length ? k.error : "string" === typeof b ? b ? b.length : 0 : b.length ? b.length : k.value
                },
                LOWER: function(b) {
                    return "string" !== typeof b ? k.value : b ? b.toLowerCase() : b
                },
                MID: function(b, a, c) {
                    a = e.parseNumber(a);
                    c = e.parseNumber(c);
                    if (e.anyIsError(a, c) || "string" !== typeof b) return c;
                    --a;
                    return b.substring(a, a + c)
                },
                NUMBERVALUE: null,
                PRONETIC: null,
                PROPER: function(b) {
                    if (void 0 === b || 0 === b.length) return k.value;
                    !0 === b && (b = "TRUE");
                    !1 === b && (b = "FALSE");
                    if (isNaN(b) && "number" === typeof b) return k.value;
                    "number" === typeof b && (b = "" + b);
                    return b.replace(/\w\S*/g, function(a) {
                        return a.charAt(0).toUpperCase() + a.substr(1).toLowerCase()
                    })
                },
                REGEXEXTRACT: function(b, a) {
                    return (b = b.match(new RegExp(a))) ? b[1 < b.length ? b.length - 1 : 0] : null
                },
                REGEXMATCH: function(b, a, c) {
                    b = b.match(new RegExp(a));
                    return c ? b : !!b
                },
                REGEXREPLACE: function(b, a, c) {
                    return b.replace(new RegExp(a), c)
                },
                REPLACE: function(b, a, c, d) {
                    a = e.parseNumber(a);
                    c = e.parseNumber(c);
                    return e.anyIsError(a, c) || "string" !== typeof b || "string" !== typeof d ? k.value : b.substr(0, a - 1) + d + b.substr(a - 1 + c)
                },
                REPT: function(b, a) {
                    a = e.parseNumber(a);
                    return a instanceof Error ? a : Array(a + 1).join(b)
                },
                RIGHT: function(b, a) {
                    a = e.parseNumber(void 0 === a ? 1 : a);
                    return a instanceof Error ? a : b ? b.substring(b.length - a) : null
                },
                SEARCH: function(b, a, c) {
                    if ("string" !== typeof b || "string" !== typeof a) return k.value;
                    c = void 0 === c ? 0 : c;
                    b = a.toLowerCase().indexOf(b.toLowerCase(), c - 1) + 1;
                    return 0 === b ? k.value : b
                },
                SPLIT: function(b, a) {
                    return b.split(a)
                },
                SUBSTITUTE: function(b, a, c, d) {
                    if (b && a && c) {
                        if (void 0 === d) return b.replace(new RegExp(a, "g"), c);
                        for (var f = 0, h = 0; 0 < b.indexOf(a, f);)
                            if (f = b.indexOf(a, f + 1), h++, h === d) return b.substring(0, f) + c + b.substring(f + a.length)
                    } else return b
                },
                T: function(b) {
                    return "string" === typeof b ? b : ""
                },
                TEXT: null,
                TRIM: function(b) {
                    return "string" !== typeof b ? k.value : b.replace(/ +/g, " ").trim()
                }
            };
            g.UNICHAR = g.CHAR;
            g.UNICODE = g.CODE;
            g.UPPER = function(b) {
                return "string" !== typeof b ? k.value : b.toUpperCase()
            };
            g.VALUE = null;
            return g
        }();
        y.stats = function() {
            var g = {
                AVEDEV: null,
                AVERAGE: function() {
                    for (var b = e.numbers(e.flatten(arguments)), a = b.length, c = 0, d = 0, f = 0; f < a; f++) c += b[f], d += 1;
                    return c / d
                },
                AVERAGEA: function() {
                    for (var b = e.flatten(arguments), a = b.length, c = 0, d = 0, f = 0; f < a; f++) {
                        var h = b[f];
                        "number" === typeof h && (c += h);
                        !0 === h && c++;
                        null !== h && d++
                    }
                    return c / d
                },
                AVERAGEIF: function(b, a, c) {
                    c = c || b;
                    b = e.flatten(b);
                    c = e.parseNumberArray(e.flatten(c));
                    if (c instanceof Error) return c;
                    for (var d = 0, f = 0, h = 0; h < b.length; h++) eval(b[h] + a) && (f += c[h], d++);
                    return f / d
                },
                AVERAGEIFS: null,
                COUNT: function() {
                    return e.numbers(e.flatten(arguments)).length
                },
                COUNTA: function() {
                    var b = e.flatten(arguments);
                    return b.length - g.COUNTBLANK(b)
                },
                COUNTIN: function(b, a) {
                    for (var c = 0, d = 0; d < b.length; d++) b[d] === a && c++;
                    return c
                },
                COUNTBLANK: function() {
                    for (var b = e.flatten(arguments), a = 0, c, d = 0; d < b.length; d++) c = b[d], null !== c && "" !== c || a++;
                    return a
                },
                COUNTIF: function() {
                    var b = e.argsToArray(arguments),
                        a = b.pop();
                    b = e.flatten(b);
                    /[<>=!]/.test(a) || (a = '=="' + a + '"');
                    for (var c = 0, d = 0; d < b.length; d++) "string" !== typeof b[d] ? eval(b[d] + a) && c++ : eval('"' + b[d] + '"' + a) && c++;
                    return c
                },
                COUNTIFS: function() {
                    for (var b =
                            e.argsToArray(arguments), a = Array(e.flatten(b[0]).length), c = 0; c < a.length; c++) a[c] = !0;
                    for (c = 0; c < b.length; c += 2) {
                        var d = e.flatten(b[c]),
                            f = b[c + 1];
                        /[<>=!]/.test(f) || (f = '=="' + f + '"');
                        for (var h = 0; h < d.length; h++) a[h] = "string" !== typeof d[h] ? a[h] && eval(d[h] + f) : a[h] && eval('"' + d[h] + '"' + f)
                    }
                    for (c = b = 0; c < a.length; c++) a[c] && b++;
                    return b
                },
                COUNTUNIQUE: function() {
                    return UNIQUE.apply(null, e.flatten(arguments)).length
                },
                FISHER: function(b) {
                    b = e.parseNumber(b);
                    return b instanceof Error ? b : Math.log((1 + b) / (1 - b)) / 2
                },
                FISHERINV: function(b) {
                    b =
                        e.parseNumber(b);
                    if (b instanceof Error) return b;
                    b = Math.exp(2 * b);
                    return (b - 1) / (b + 1)
                },
                FREQUENCY: function(b, a) {
                    b = e.parseNumberArray(e.flatten(b));
                    a = e.parseNumberArray(e.flatten(a));
                    if (e.anyIsError(b, a)) return k.value;
                    for (var c = b.length, d = a.length, f = [], h = 0; h <= d; h++)
                        for (var l = f[h] = 0; l < c; l++) 0 === h ? b[l] <= a[0] && (f[0] += 1) : h < d ? b[l] > a[h - 1] && b[l] <= a[h] && (f[h] += 1) : h === d && b[l] > a[d - 1] && (f[d] += 1);
                    return f
                },
                LARGE: function(b, a) {
                    b = e.parseNumberArray(e.flatten(b));
                    a = e.parseNumber(a);
                    return e.anyIsError(b, a) ? b : b.sort(function(c,
                        d) {
                        return d - c
                    })[a - 1]
                },
                MAX: function() {
                    var b = e.numbers(e.flatten(arguments));
                    return 0 === b.length ? 0 : Math.max.apply(Math, b)
                },
                MAXA: function() {
                    var b = e.arrayValuesToNumbers(e.flatten(arguments));
                    return 0 === b.length ? 0 : Math.max.apply(Math, b)
                },
                MIN: function() {
                    var b = e.numbers(e.flatten(arguments));
                    return 0 === b.length ? 0 : Math.min.apply(Math, b)
                },
                MINA: function() {
                    var b = e.arrayValuesToNumbers(e.flatten(arguments));
                    return 0 === b.length ? 0 : Math.min.apply(Math, b)
                },
                MODE: {}
            };
            g.MODE.MULT = function() {
                var b = e.parseNumberArray(e.flatten(arguments));
                if (b instanceof Error) return b;
                for (var a = b.length, c = {}, d = [], f = 0, h, l = 0; l < a; l++) h = b[l], c[h] = c[h] ? c[h] + 1 : 1, c[h] > f && (f = c[h], d = []), c[h] === f && (d[d.length] = h);
                return d
            };
            g.MODE.SNGL = function() {
                var b = e.parseNumberArray(e.flatten(arguments));
                return b instanceof Error ? b : g.MODE.MULT(b).sort(function(a, c) {
                    return a - c
                })[0]
            };
            g.PERCENTILE = {};
            g.PERCENTILE.EXC = function(b, a) {
                b = e.parseNumberArray(e.flatten(b));
                a = e.parseNumber(a);
                if (e.anyIsError(b, a)) return k.value;
                b = b.sort(function(d, f) {
                    return d - f
                });
                var c = b.length;
                if (a <
                    1 / (c + 1) || a > 1 - 1 / (c + 1)) return k.num;
                a = a * (c + 1) - 1;
                c = Math.floor(a);
                return e.cleanFloat(a === c ? b[a] : b[c] + (a - c) * (b[c + 1] - b[c]))
            };
            g.PERCENTILE.INC = function(b, a) {
                b = e.parseNumberArray(e.flatten(b));
                a = e.parseNumber(a);
                if (e.anyIsError(b, a)) return k.value;
                b = b.sort(function(d, f) {
                    return d - f
                });
                a *= b.length - 1;
                var c = Math.floor(a);
                return e.cleanFloat(a === c ? b[a] : b[c] + (a - c) * (b[c + 1] - b[c]))
            };
            g.PERCENTRANK = {};
            g.PERCENTRANK.EXC = function(b, a, c) {
                c = void 0 === c ? 3 : c;
                b = e.parseNumberArray(e.flatten(b));
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                if (e.anyIsError(b, a, c)) return k.value;
                b = b.sort(function(p, q) {
                    return p - q
                });
                var d = UNIQUE.apply(null, b),
                    f = b.length,
                    h = d.length;
                c = Math.pow(10, c);
                for (var l = 0, m = !1, n = 0; !m && n < h;) a === d[n] ? (l = (b.indexOf(d[n]) + 1) / (f + 1), m = !0) : a >= d[n] && (a < d[n + 1] || n === h - 1) && (l = (b.indexOf(d[n]) + 1 + (a - d[n]) / (d[n + 1] - d[n])) / (f + 1), m = !0), n++;
                return Math.floor(l * c) / c
            };
            g.PERCENTRANK.INC = function(b, a, c) {
                c = void 0 === c ? 3 : c;
                b = e.parseNumberArray(e.flatten(b));
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                if (e.anyIsError(b, a, c)) return k.value;
                b = b.sort(function(p,
                    q) {
                    return p - q
                });
                var d = UNIQUE.apply(null, b),
                    f = b.length,
                    h = d.length;
                c = Math.pow(10, c);
                for (var l = 0, m = !1, n = 0; !m && n < h;) a === d[n] ? (l = b.indexOf(d[n]) / (f - 1), m = !0) : a >= d[n] && (a < d[n + 1] || n === h - 1) && (l = (b.indexOf(d[n]) + (a - d[n]) / (d[n + 1] - d[n])) / (f - 1), m = !0), n++;
                return Math.floor(l * c) / c
            };
            g.PERMUT = function(b, a) {
                b = e.parseNumber(b);
                a = e.parseNumber(a);
                return e.anyIsError(b, a) ? k.value : FACT(b) / FACT(b - a)
            };
            g.PERMUTATIONA = function(b, a) {
                b = e.parseNumber(b);
                a = e.parseNumber(a);
                return e.anyIsError(b, a) ? k.value : Math.pow(b, a)
            };
            g.PHI =
                function(b) {
                    b = e.parseNumber(b);
                    return b instanceof Error ? k.value : Math.exp(-.5 * b * b) / 2.5066282746310002
                };
            g.PROB = function(b, a, c, d) {
                if (void 0 === c) return 0;
                d = void 0 === d ? c : d;
                b = e.parseNumberArray(e.flatten(b));
                a = e.parseNumberArray(e.flatten(a));
                c = e.parseNumber(c);
                d = e.parseNumber(d);
                if (e.anyIsError(b, a, c, d)) return k.value;
                if (c === d) return 0 <= b.indexOf(c) ? a[b.indexOf(c)] : 0;
                for (var f = b.sort(function(n, p) {
                        return n - p
                    }), h = f.length, l = 0, m = 0; m < h; m++) f[m] >= c && f[m] <= d && (l += a[b.indexOf(f[m])]);
                return l
            };
            g.QUARTILE = {};
            g.QUARTILE.EXC = function(b, a) {
                b = e.parseNumberArray(e.flatten(b));
                a = e.parseNumber(a);
                if (e.anyIsError(b, a)) return k.value;
                switch (a) {
                    case 1:
                        return g.PERCENTILE.EXC(b, .25);
                    case 2:
                        return g.PERCENTILE.EXC(b, .5);
                    case 3:
                        return g.PERCENTILE.EXC(b, .75);
                    default:
                        return k.num
                }
            };
            g.QUARTILE.INC = function(b, a) {
                b = e.parseNumberArray(e.flatten(b));
                a = e.parseNumber(a);
                if (e.anyIsError(b, a)) return k.value;
                switch (a) {
                    case 1:
                        return g.PERCENTILE.INC(b, .25);
                    case 2:
                        return g.PERCENTILE.INC(b, .5);
                    case 3:
                        return g.PERCENTILE.INC(b,
                            .75);
                    default:
                        return k.num
                }
            };
            g.RANK = {};
            g.RANK.AVG = function(b, a, c) {
                b = e.parseNumber(b);
                a = e.parseNumberArray(e.flatten(a));
                if (e.anyIsError(b, a)) return k.value;
                a = e.flatten(a);
                a = a.sort(c ? function(h, l) {
                    return h - l
                } : function(h, l) {
                    return l - h
                });
                c = a.length;
                for (var d = 0, f = 0; f < c; f++) a[f] === b && d++;
                return 1 < d ? (2 * a.indexOf(b) + d + 1) / 2 : a.indexOf(b) + 1
            };
            g.RANK.EQ = function(b, a, c) {
                b = e.parseNumber(b);
                a = e.parseNumberArray(e.flatten(a));
                if (e.anyIsError(b, a)) return k.value;
                a = a.sort(c ? function(d, f) {
                    return d - f
                } : function(d, f) {
                    return f -
                        d
                });
                return a.indexOf(b) + 1
            };
            g.RSQ = function(b, a) {
                b = e.parseNumberArray(e.flatten(b));
                a = e.parseNumberArray(e.flatten(a));
                return e.anyIsError(b, a) ? k.value : Math.pow(g.PEARSON(b, a), 2)
            };
            g.SMALL = function(b, a) {
                b = e.parseNumberArray(e.flatten(b));
                a = e.parseNumber(a);
                return e.anyIsError(b, a) ? b : b.sort(function(c, d) {
                    return c - d
                })[a - 1]
            };
            g.STANDARDIZE = function(b, a, c) {
                b = e.parseNumber(b);
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(b, a, c) ? k.value : (b - a) / c
            };
            g.STDEV = {};
            g.STDEV.P = function() {
                var b = g.VAR.P.apply(this,
                    arguments);
                return Math.sqrt(b)
            };
            g.STDEV.S = function() {
                var b = g.VAR.S.apply(this, arguments);
                return Math.sqrt(b)
            };
            g.STDEVA = function() {
                var b = g.VARA.apply(this, arguments);
                return Math.sqrt(b)
            };
            g.STDEVPA = function() {
                var b = g.VARPA.apply(this, arguments);
                return Math.sqrt(b)
            };
            g.VAR = {};
            g.VAR.P = function() {
                for (var b = e.numbers(e.flatten(arguments)), a = b.length, c = 0, d = g.AVERAGE(b), f = 0; f < a; f++) c += Math.pow(b[f] - d, 2);
                return c / a
            };
            g.VAR.S = function() {
                for (var b = e.numbers(e.flatten(arguments)), a = b.length, c = 0, d = g.AVERAGE(b),
                        f = 0; f < a; f++) c += Math.pow(b[f] - d, 2);
                return c / (a - 1)
            };
            g.VARA = function() {
                for (var b = e.flatten(arguments), a = b.length, c = 0, d = 0, f = g.AVERAGEA(b), h = 0; h < a; h++) {
                    var l = b[h];
                    c = "number" === typeof l ? c + Math.pow(l - f, 2) : !0 === l ? c + Math.pow(1 - f, 2) : c + Math.pow(0 - f, 2);
                    null !== l && d++
                }
                return c / (d - 1)
            };
            g.VARPA = function() {
                for (var b = e.flatten(arguments), a = b.length, c = 0, d = 0, f = g.AVERAGEA(b), h = 0; h < a; h++) {
                    var l = b[h];
                    c = "number" === typeof l ? c + Math.pow(l - f, 2) : !0 === l ? c + Math.pow(1 - f, 2) : c + Math.pow(0 - f, 2);
                    null !== l && d++
                }
                return c / d
            };
            g.WEIBULL = {};
            g.WEIBULL.DIST = function(b, a, c, d) {
                b = e.parseNumber(b);
                a = e.parseNumber(a);
                c = e.parseNumber(c);
                return e.anyIsError(b, a, c) ? k.value : d ? 1 - Math.exp(-Math.pow(b / c, a)) : Math.pow(b, a - 1) * Math.exp(-Math.pow(b / c, a)) * a / Math.pow(c, a)
            };
            g.Z = {};
            g.Z.TEST = function(b, a, c) {
                b = e.parseNumberArray(e.flatten(b));
                a = e.parseNumber(a);
                if (e.anyIsError(b, a)) return k.value;
                c = c || g.STDEV.S(b);
                var d = b.length;
                return 1 - g.NORM.S.DIST((g.AVERAGE(b) - a) / (c / Math.sqrt(d)), !0)
            };
            return g
        }();
        y.utils = function() {
            return {
                PROGRESS: function(g, b) {
                    return '<div style="width:' +
                        (g ? g : "0") + "%;height:4px;background-color:" + (b ? b : "red") + ';margin-top:1px;"></div>'
                },
                RATING: function(g) {
                    for (var b = '<div class="jrating">', a = 0; 5 > a; a++) b = a < g ? b + '<div class="jrating-selected"></div>' : b + "<div></div>";
                    return b + "</div>"
                }
            }
        }();
        for (var H = 0; H < Object.keys(y).length; H++)
            for (var C = y[Object.keys(y)[H]], z = Object.keys(C), x = 0; x < z.length; x++)
                if (C[z[x]])
                    if ("function" == typeof C[z[x]] || "object" == typeof C[z[x]]) {
                        if (w[z[x]] = C[z[x]], w[z[x]].toString = function() {
                                return "#ERROR"
                            }, "object" == typeof C[z[x]])
                            for (var J =
                                    Object.keys(C[z[x]]), I = 0; I < J.length; I++) w[z[x]][J[I]].toString = function() {
                                return "#ERROR"
                            }
                    } else w[z[x]] = function() {
                        return z[x] + "Not implemented"
                    };
        else w[z[x]] = function() {
            return z[x] + "Not implemented"
        };
        var F = null,
            D = null,
            B = null;
        w.TABLE = function() {
            return B
        };
        w.COLUMN = w.COL = function() {
            B.tracking && B.tracking.push(A.getColumnNameFromCoords(parseInt(F), parseInt(D)));
            return parseInt(F) + 1
        };
        w.ROW = function() {
            B.tracking && B.tracking.push(A.getColumnNameFromCoords(parseInt(F), parseInt(D)));
            return parseInt(D) + 1
        };
        w.CELL = function() {
            return A.getColumnNameFromCoords(F, D)
        };
        w.VALUE = function(g, b, a) {
            return B.getValueFromCoords(parseInt(g) - 1, parseInt(b) - 1, a)
        };
        w.THISROWCELL = function(g) {
            return B.getValueFromCoords(parseInt(g) - 1, parseInt(D))
        };
        var K = function(g, b) {
                for (var a = 0; a < g.length; a++) {
                    var c = A.getTokensFromRange(g[a]);
                    b = b.replace(g[a], "[" + c.join(",") + "]")
                }
                return b
            },
            A = function(g, b, a, c, d) {
                B = d;
                F = a;
                D = c;
                c = "";
                d = Object.keys(b);
                if (d.length) {
                    var f = {};
                    for (a = 0; a < d.length; a++)
                        if (h = d[a].replace(/!/g, "."), 0 < h.indexOf(".")) {
                            var h =
                                h.split(".");
                            f[h[0]] = {}
                        } h = Object.keys(f);
                    for (a = 0; a < h.length; a++) c += "var " + h[a] + " = {};";
                    for (a = 0; a < d.length; a++) {
                        h = d[a].replace(/!/g, ".");
                        if (f = null !== b[d[a]]) f = b[d[a]], f = !(!isNaN(f) && null !== f && "" !== f);
                        f && (f = b[d[a]].match(/(('.*?'!)|(\w*!))?(\$?[A-Z]+\$?[0-9]*):(\$?[A-Z]+\$?[0-9]*)?/g)) && f.length && (b[d[a]] = K(f, b[d[a]]));
                        c = 0 < h.indexOf(".") ? c + (h + " = " + b[d[a]] + ";\n") : c + ("var " + h + " = " + b[d[a]] + ";\n")
                    }
                }
                g = g.replace(/\$/g, "");
                g = g.replace(/!/g, ".");
                b = "";
                a = 0;
                d = ["=", "!", ">", "<"];
                for (h = 0; h < g.length; h++) '"' == g[h] &&
                    (a = 0 == a ? 1 : 0), 1 == a ? b += g[h] : (b += g[h].toUpperCase(), 0 < h && "=" == g[h] && -1 == d.indexOf(g[h - 1]) && -1 == d.indexOf(g[h + 1]) && (b += "="));
                b = b.replace(/\^/g, "**");
                b = b.replace(/<>/g, "!=");
                b = b.replace(/&/g, "+");
                g = b = b.replace(/\$/g, "");
                (f = g.match(/(('.*?'!)|(\w*!))?(\$?[A-Z]+\$?[0-9]*):(\$?[A-Z]+\$?[0-9]*)?/g)) && f.length && (g = K(f, g));
                c = (new Function(c + "; return " + g))();
                null === c && (c = 0);
                return c
            };
        A.getColumnNameFromCoords = function(g, b) {
            g = parseInt(g);
            var a = "";
            701 < g ? (a += String.fromCharCode(64 + parseInt(g / 676)), a += String.fromCharCode(64 +
                parseInt(g % 676 / 26))) : 25 < g && (a += String.fromCharCode(64 + parseInt(g / 26)));
            a += String.fromCharCode(65 + g % 26);
            return a + (parseInt(b) + 1)
        };
        A.getCoordsFromColumnName = function(g) {
            var b = /^[a-zA-Z]+/.exec(g);
            if (b) {
                for (var a = 0, c = 0; c < b[0].length; c++) a += parseInt(b[0].charCodeAt(c) - 64) * Math.pow(26, b[0].length - 1 - c);
                a--;
                0 > a && (a = 0);
                g = parseInt(/[0-9]+$/.exec(g)) || null;
                0 < g && g--;
                return [a, g]
            }
        };
        A.getRangeFromTokens = function(g) {
            g = g.filter(function(d) {
                return "#REF!" != d
            });
            for (var b = "", a = "", c = 0; c < g.length; c++) 0 <= g[c].indexOf(".") ?
                b = "." : 0 <= g[c].indexOf("!") && (b = "!"), b && (a = g[c].split(b), g[c] = a[1], a = a[0] + b);
            g.sort(function(d, f) {
                d = Helpers.getCoordsFromColumnName(d);
                f = Helpers.getCoordsFromColumnName(f);
                return d[1] > f[1] ? 1 : d[1] < f[1] ? -1 : d[0] > f[0] ? 1 : d[0] < f[0] ? -1 : 0
            });
            return g.length ? a + (g[0] + ":" + g[g.length - 1]) : "#REF!"
        };
        A.getTokensFromRange = function(g) {
            if (0 < g.indexOf(".")) {
                var b = g.split(".");
                g = b[1];
                b = b[0] + "."
            } else 0 < g.indexOf("!") ? (b = g.split("!"), g = b[1], b = b[0] + "!") : b = "";
            g = g.split(":");
            var a = A.getCoordsFromColumnName(g[0]),
                c = A.getCoordsFromColumnName(g[1]);
            if (a[0] <= c[0]) {
                g = a[0];
                var d = c[0]
            } else g = c[0], d = a[0];
            if (null === a[1] && null == c[1])
                for (var f = null, h = null, l = Object.keys(vars), m = 0; m < l.length; m++) {
                    var n = A.getCoordsFromColumnName(l[m]);
                    n[0] === a[0] && (null === f || n[1] < f) && (f = n[1]);
                    n[0] === c[0] && (null === h || n[1] > h) && (h = n[1])
                } else a[1] <= c[1] ? (f = a[1], h = c[1]) : (f = c[1], h = a[1]);
            for (a = []; f <= h; f++) {
                c = [];
                for (m = g; m <= d; m++) c.push(b + A.getColumnNameFromCoords(m, f));
                a.push(c)
            }
            return a
        };
        A.setFormula = function(g) {
            for (var b = Object.keys(g), a = 0; a < b.length; a++) "function" == typeof g[b[a]] &&
                (w[b[a]] = g[b[a]])
        };
        return A
    };
    return "undefined" !== typeof window ? G(window) : null
});

if (!jSuites && typeof(require) === 'function') {
    var jSuites = require('jsuites');
}

var $jscomp = $jscomp || {};
$jscomp.scope = {};
$jscomp.arrayIteratorImpl = function(D) {
    var I = 0;
    return function() {
        return I < D.length ? {
            done: !1,
            value: D[I++]
        } : {
            done: !0
        }
    }
};
$jscomp.arrayIterator = function(D) {
    return {
        next: $jscomp.arrayIteratorImpl(D)
    }
};
$jscomp.ASSUME_ES5 = !1;
$jscomp.ASSUME_NO_NATIVE_MAP = !1;
$jscomp.ASSUME_NO_NATIVE_SET = !1;
$jscomp.SIMPLE_FROUND_POLYFILL = !1;
$jscomp.ISOLATE_POLYFILLS = !1;
$jscomp.FORCE_POLYFILL_PROMISE = !1;
$jscomp.FORCE_POLYFILL_PROMISE_WHEN_NO_UNHANDLED_REJECTION = !1;
$jscomp.defineProperty = $jscomp.ASSUME_ES5 || "function" == typeof Object.defineProperties ? Object.defineProperty : function(D, I, M) {
    if (D == Array.prototype || D == Object.prototype) return D;
    D[I] = M.value;
    return D
};
$jscomp.getGlobal = function(D) {
    D = ["object" == typeof globalThis && globalThis, D, "object" == typeof window && window, "object" == typeof self && self, "object" == typeof global && global];
    for (var I = 0; I < D.length; ++I) {
        var M = D[I];
        if (M && M.Math == Math) return M
    }
    throw Error("Cannot find global object");
};
$jscomp.global = $jscomp.getGlobal(this);
$jscomp.IS_SYMBOL_NATIVE = "function" === typeof Symbol && "symbol" === typeof Symbol("x");
$jscomp.TRUST_ES6_POLYFILLS = !$jscomp.ISOLATE_POLYFILLS || $jscomp.IS_SYMBOL_NATIVE;
$jscomp.polyfills = {};
$jscomp.propertyToPolyfillSymbol = {};
$jscomp.POLYFILL_PREFIX = "$jscp$";
var $jscomp$lookupPolyfilledValue = function(D, I) {
    var M = $jscomp.propertyToPolyfillSymbol[I];
    if (null == M) return D[I];
    M = D[M];
    return void 0 !== M ? M : D[I]
};
$jscomp.polyfill = function(D, I, M, P) {
    I && ($jscomp.ISOLATE_POLYFILLS ? $jscomp.polyfillIsolated(D, I, M, P) : $jscomp.polyfillUnisolated(D, I, M, P))
};
$jscomp.polyfillUnisolated = function(D, I, M, P) {
    M = $jscomp.global;
    D = D.split(".");
    for (P = 0; P < D.length - 1; P++) {
        var W = D[P];
        if (!(W in M)) return;
        M = M[W]
    }
    D = D[D.length - 1];
    P = M[D];
    I = I(P);
    I != P && null != I && $jscomp.defineProperty(M, D, {
        configurable: !0,
        writable: !0,
        value: I
    })
};
$jscomp.polyfillIsolated = function(D, I, M, P) {
    var W = D.split(".");
    D = 1 === W.length;
    P = W[0];
    P = !D && P in $jscomp.polyfills ? $jscomp.polyfills : $jscomp.global;
    for (var Z = 0; Z < W.length - 1; Z++) {
        var ra = W[Z];
        if (!(ra in P)) return;
        P = P[ra]
    }
    W = W[W.length - 1];
    M = $jscomp.IS_SYMBOL_NATIVE && "es6" === M ? P[W] : null;
    I = I(M);
    null != I && (D ? $jscomp.defineProperty($jscomp.polyfills, W, {
        configurable: !0,
        writable: !0,
        value: I
    }) : I !== M && (void 0 === $jscomp.propertyToPolyfillSymbol[W] && (M = 1E9 * Math.random() >>> 0, $jscomp.propertyToPolyfillSymbol[W] = $jscomp.IS_SYMBOL_NATIVE ?
        $jscomp.global.Symbol(W) : $jscomp.POLYFILL_PREFIX + M + "$" + W), $jscomp.defineProperty(P, $jscomp.propertyToPolyfillSymbol[W], {
        configurable: !0,
        writable: !0,
        value: I
    })))
};
$jscomp.initSymbol = function() {};
$jscomp.iteratorPrototype = function(D) {
    D = {
        next: D
    };
    D[Symbol.iterator] = function() {
        return this
    };
    return D
};
$jscomp.iteratorFromArray = function(D, I) {
    D instanceof String && (D += "");
    var M = 0,
        P = !1,
        W = {
            next: function() {
                if (!P && M < D.length) {
                    var Z = M++;
                    return {
                        value: I(Z, D[Z]),
                        done: !1
                    }
                }
                P = !0;
                return {
                    done: !0,
                    value: void 0
                }
            }
        };
    W[Symbol.iterator] = function() {
        return W
    };
    return W
};
$jscomp.polyfill("Array.prototype.values", function(D) {
    return D ? D : function() {
        return $jscomp.iteratorFromArray(this, function(I, M) {
            return M
        })
    }
}, "es8", "es3");
$jscomp.polyfill("Object.fromEntries", function(D) {
    return D ? D : function(I) {
        var M = {};
        if (!(Symbol.iterator in I)) throw new TypeError("" + I + " is not iterable");
        I = I[Symbol.iterator].call(I);
        for (var P = I.next(); !P.done; P = I.next()) {
            P = P.value;
            if (Object(P) !== P) throw new TypeError("iterable for fromEntries should yield objects");
            M[P[0]] = P[1]
        }
        return M
    }
}, "es_2019", "es3");
var _$_669e = ["use strict", "{", "}", "", "9.3.0", "Base", "https://jspreadsheet.com", "Unlicensed", "Jspreadsheet Pro\r\n", " edition ", "\r\n", "set", "&", "&lt;", "g", "<", "&gt;", ">", "=", "!", '"', "undefined", "number", "fullscreen", "block", "string", "readonly", "jss_rotate", "span", "rotate(", "deg)", "jss_hidden_index", "div", "right", "10px", "pointer", "/v9", ",", "localhost", ".", "v9", " (Expired)", "Licensed to: ", "Premium", "License required", "A valid license is required", "This version is not included on the scope of this license",
    "Your license is expired. Please renew your certificate.", "jspreadsheet", "img", "/jspreadsheet/", "logo.png", "none", "status", "edition", "scope", "get", "px", "jss_frozen", "data-y", "scroll", "wheel", "goto", "data-x", "jss_scroll", "jss_hidden", "jss_content_overflow", "onresize", "Jspreadsheet Pro", "text", "autocomplete", "dropdown", "GET", "json", "Jspreadsheet Alert: secureFormulas is set to false.", "setConfig", "insertRow", "deleteRow", "insertColumn", "deleteColumn", "moveRow", "moveColumn", "hideRow", "showRow", "hideColumn",
    "showColumn", "setMerge", "removeMerge", "setStyle", "resetStyle", "setWidth", "setHeight", "setHeader", "setComments", "setFreezeColumns", "setFreezeRows", "orderBy", "setValue", "setProperty", "updateProperty", "renameWorksheet", "moveWorksheet", "setDefinedNames", "setRowGroup", "resetRowGroup", "openRowGroup", "closeRowGroup", "onundo", "onredo", "object", "function", "contextMenu", "toolbar", "onevent", "Jspreadsheet: cursor not in the viewport", "copying", "keydown", "oneditionstart", "oncreateeditor", "hidden", "jss_focus",
    "jss_formula", "oneditionend", "data-mask", "data-value", "inputmode", "0", "style", "jss_input", "contentEditable", "oninput", "setRowId", "onchangerowid", "setBorder", "resetBorders", "updateCells", "setFormula", "updateData", "onbeforesave", "POST", "Authorization", "Bearer ", "onbeforesend", "disconected", ": ", "Sorry, something went wrong, refreshing your spreadsheet...", "_top", "onsave", "disconnected", "persistence", "?", "?worksheet=", "&worksheet=", "add", "remove", "jss_cursor", "selected", "main", "onblur", "onfocus", "onselection",
    " ", "'", "'!", ":", "jss_border", "cloning", "jss_border_", "15", "jss_resize_corner", "data-action", "ontouchstart", "moveLeft", "0px", "-6px", "moveTop", "moveRight", "moveDown", "-2000px", "transparent", "left", "bottom", "top", "tr", "jss_group_container", "td", "&nbsp;", ".jss_group", "jss_group", "+", "oncreaterowgroup", "i", "material-icons", "onopenrowgroup", "oncloserowgroup", "ondestroyrowgroup", "width", "Jspreadsheet: This row is part of a merged cell", "Jspreadsheet: Please clear your search before perform this action",
    "jss_dragging", "1px", "jss_helper_row", "Jspreadsheet: This column is part of a merged cell.", "jss_helper_col", "No cell selected", "Invalid merged properties", "onmerge", "colspan", "rowspan", "Column name", "data-title", "onchangeheader", "jss_header", "col", "oncreatecolumn", "jss_filters_icon", "onrenderfootercell", "setFooter", "onchangefooter", "setFooterValue", "onchangefootervalue", "resetFooter", "setNestedHeaders", "onchangenested", "jss_nested", "resetNestedHeaders", "setNestedCell", "onchangenestedcell", "center",
    "data-nested-x", "data-nested-y", "align", "calendar", '""', "\t", "oncopy", "copy", "color: initial", "onbeforepaste", "JSS: For better results, please allow pagination or tableOverflow.", "onpaste", "Export not allowed", "\ufeff", "text/csv;charset=utf-8;", ".csv", "a", "download", "Defined names are only available on the Premium edition.", "onchangedefinednames", "#REF!", "#1a237e", "#b71c1c", "#880e4f", "#1b5e20", "#ff6f00", "-", "/", "*", "^", "(", "[", "|", ";", "data-token", "picker", "jss_picker", "click", "active", "focus", "input",
    "jss_object", "contenteditable", "blur", "last", '<b class="', '" data-token="', '">', "</b>", "data-single", "true", "b", ")", "]", "var ", " = {};", " = ", ";\n", "; return ", "onbeforeformula", "#LOOP", "Reference loop detected", "#SELF", "Self Reference detected", "#ERROR", "Table does not exist or is not loaded yet", "CHART(", "#SPILL!", "jss_number", "$", "Something went wrong", "onafterchanges", "checkbox", "autonumber", "email", "url", "color", "progressbar", "rating", "radio", "onformulachain", "onbeforechange", "onchange", "oncreatecell",
    "values", "mergeCells", "title", "jss_comments", "jss_unlocked", "jss_zebra", "jss_wrap", "jss_hidden_col", "onchangeproperty", "onchangerowvisibility", "jss_row", "oncreaterow", "Jspreadsheet: Invalid origin or destination", "onmoverow", "Jspreadsheet: Insert row is not enabled in the table configuration", "onbeforeinsertrow", "onbeforeinsertrow returned false", "oninsertrow", "Jspreadsheet: Delete row is not enabled in the table configuration", "onbeforedeleterow", "onbeforedeleterow returned false", "meta", "cells",
    "comments", "ondeleterow", "onchangecolumnvisibility", "Jspreadsheet: Invalid origin", "Jspreadsheet: Invalid destination", "onmovecolumn", "Jspreadsheet: Insert column is not enabled in the table configuration", "onbeforeinsertcolumn", "onbeforeinsertcolumn returned false", "record", "col_", "oninsertcolumn", "Jspreadsheet: Delete column is not enabled in the table configuration", "Jspreadsheet: it is not possible to delete the last column", "onbeforedeletecolumn", "onbeforedeletecolumn returned false", "ondeletecolumn",
    "; ", "onchangestyle", "No cell provided", "onresetstyle", "onbeforecomments", "oncomments", "setMeta", "onchangemeta", "resetMeta", "onresetmeta", "arrow-down", "percent", "numeric", "onbeforesort", "This action will destroy any existing merged cells. Are you sure?", "onsort", "arrow-up", "jss_pagination", "Jspreadsheet: No pagination defined", "Jspreadsheet: pagination not defined", "onbeforechangepage", "onchangepage", "No records found", "jss_page", "jss_page_selected", "Showing page {0} of {1} entries", "jss_frozen_row",
    "Frozen rows cannot be larger than the viewport.", "Frozen rows cannot be over merged rows.", "jss_freeze_control_row", "Frozen columns cannot be larger than the viewport.", "Frozen columns cannot be over merged columns.", "jss_freeze_control_col", "jss_content", "selectstart", "table", "thead", "tbody", "tfoot", "jss_nogridline", "colgroup", "50", "jss_selectall", "JSS: Freeze controls are only available on the Premium edition.", "jss", "cellpadding", "cellspacing", "unselectable", "yes", "jss_", "jss_overflow", "jss_locked",
    "jss_corner", "on", "onselectstart", "return false", "jss_content_pagination", "Frozen columns cannot be greater than the number of available columns", "rows", "onbeforechangereferences", "onchangereferences", "jss_container", "jss_toolbar", "jss_helper", "jss_loading", "textarea", "jss_textarea", "-1", "beforeinit", "Sheet", "Worksheet clash name: ", ". It is highly recommended to define a unique worksheetName on the initialization.", "jss_worksheet", "init", "To improve the performance, please consider tableOverflow or pagination.",
    "resize", "touchstart", "touchmove", "touchend", "auto", "onresizecolumn", "onresizerow", "label", "jss_filters_active", "Jspreadsheet: the filter is not enabled.", "onopenfilter", "Blanks", "onbeforefilter", "onfilter", "gi", "jss_filters_options", "INPUT", "Search", "jss_filters_search", "button", "Ok", "jss_filters_apply", "jss_filters", '<input type="checkbox"> ', "Add current selection to filter", '<input type="checkbox"> (', "Select all", "select", "jss_pagination_dropdown", "option", "Show", "entries", "jss_search", "jss_searching",
    "onsearchrow", "onsearchstart", "onbeforesearch", "onsearch", "jss_search_container", "jss_contextmenu", "cell", "row", "header", "Mac", "&#8984;", "Ctrl", "tabs", "Rename this worksheet", "Delete this worksheet", "Are you sure?", "line", "nested", "Rename this cell", "Cut", "content_cut", " + X", "Copy", "content_copy", " + C", "Paste", "content_paste", " + V", "Insert a new column before", "Insert a new column after", "Delete selected columns", "Rename this column", "Create a new row", "Order ascending", "Order descending", "Hide",
    "Insert a new row before", "Insert a new row after", "Delete selected rows", "Edit notes", "Add notes", "notes", "Notes", "Clear notes", "selectall", "Show all", "Save as", "save", " + S", "About", "info", "with-toolbar", "jtoolbar-disabled", "undo", "redo", "divisor", "120px", "Default", "Verdana", "Arial", "Courier New", '<span style="font-family:', "</span>", "font-family", "48px", "format_size", "x-small", "small", "medium", "large", "x-large", '<span style="font-size:', "font-size", "justify", '<i class="material-icons">format_align_',
    "</i>", "text-align", "format_bold", "fontWeight", "bold", "font-weight", "jss_toolbar_selected", "format_color_text", "format_color_fill", "background-color", "extended", "middle", "vertical_align_top", "vertical_align_center", "vertical_align_bottom", '<i class="material-icons">', "vertical-align", "web", "Merge the selected cells", "border_all", "border_outer", "border_inner", "border_horizontal", "border_vertical", "border_left", "border_top", "border_right", "border_bottom", "border_clear", "black", "border-left: ", "px solid ",
    "border-left: ; ", "border-right: ", "border-right: ; ", "border-top: ", "border-top: ; ", "border-bottom: ", "border-bottom: ; ", "color_lens", '<div style="height: ', 'px; width: 50px; background-color: black;"></div>', "80px", "1", "fullscreen_exit", "Toggle Fullscreen", "search", "search_off", "Toggle Search", "onbeforecreateworksheet", "There is one existing worksheet with the same name.", "createWorksheet", "oncreateworksheet", "deleteWorksheet", "ondeleteworksheet", "It was not possible to rename worksheet due conflict name",
    "onrenameworksheet", "onmoveworksheet", "onopenworksheet", "TABLE", "Element is not a table", "colgroup > col", "data-celltype", "readOnly", "name", "id", ":scope > thead > tr", ":scope > tr, :scope > tbody > tr", "data-formula", "class", "styleBold", "; font-weight:bold;", "font-weight:bold;", "tfoot tr", "\r", "\n", "jss_ignore", "THEAD", "TBODY", "TFOOT", "header-group", "jtabs-headers", "filters", "footer", "jtoolbar", "DIV", "html", "jss_nowrap", "insertHTML", "<br/>\n", "Process", "move", "col-resize", "grab", "row-resize", "A",
    "mailto:", "_blank", "onclick", "jclose", "jss_dropdown", "DOM el not found", '<table class="jss"><tbody><tr><td>A</td></tr></tbody></table>', "Cloud extension not loaded.", "onload", "jss_dialog", "@", "decimal", "data-locale", "%", "jss_percent", "[Red]", "red", "jss_notes", "145px", "tag", "jss_dropdown_tags", "multiple", "image", "<div class='jss_dropdown_icon' style='background-color: ", "'></div>", "orange", "<div class='jss_dropdown_tag' style='background-color: ", "'>", "</div>", "timepicker", "YYYY-MM-DD", "square", "100px",
    "false", "FALSE", "TRUE", "jss_progressbar", "min", "max", "jss_rating", " stars", "star", "href", "IMG", "round", "100%", "jss_image", "tabindex", "jpanel", "delete", "jfile", '<img src="', '" alt="">', "jss_richtext", "jss_tags", "40px", "#NOTFOUND", "<img src='", "' class='round small' />", "Jspreadsheet: worksheet not found "
];
(function(D, I) {
    typeof exports === _$_669e[109] && typeof module !== _$_669e[21] ? module.exports = I() : typeof define === _$_669e[110] && define.amd ? define(I) : D.jspreadsheet = D.jexcel = I()
})(this, function() {
    _$_669e[0];
    var D = function(h, c) {
            h = jSuites.translate(h);
            if (c && c.length)
                for (var b = 0; b < c.length; b++) h = h.replace(_$_669e[1] + b + _$_669e[2], c[b]);
            return h
        },
        I = function(h, c) {
            Array.isArray(h) && (h = h[0]);
            h = h.replace(new RegExp(/'/g), _$_669e[3]).toUpperCase();
            if (void 0 === c) return window[h] ? window[h] : null;
            window[h] = c
        },
        M = function() {
            var h = {
                version: _$_669e[4],
                edition: _$_669e[5],
                host: _$_669e[6],
                license: _$_669e[7],
                print: function() {
                    return [_$_669e[8] + this.edition + _$_669e[9] + this.version + _$_669e[10] + this.host + _$_669e[10] + this.license]
                }
            };
            return function() {
                return h
            }
        }(),
        P = function(h) {
            var c = parseInt(h.y);
            if (h.id) var b = parseInt(h.id);
            else h.value && (b = parseInt(h.value));
            0 < b && (this.rows[c].id = b, za[_$_669e[11]].call(this, b))
        },
        W = function(h, c) {
            var b = _$_669e[3],
                a = 0; - 1 !== h.indexOf(_$_669e[12]) && (h = h.replace(new RegExp(_$_669e[13], _$_669e[14]), _$_669e[15]),
                h = h.replace(new RegExp(_$_669e[16], _$_669e[14]), _$_669e[17]));
            for (var d = [_$_669e[18], _$_669e[19], _$_669e[17], _$_669e[15]], e = 0; e < h.length; e++) h[e] == _$_669e[20] && (a = 0 == a ? 1 : 0), 1 == a ? b += h[e] : (b += h[e].toUpperCase(), 1 == c && 0 < e && h[e] == _$_669e[18] && -1 == d.indexOf(h[e - 1]) && -1 == d.indexOf(h[e + 1]) && (b += _$_669e[18]));
            return b
        },
        Z = function(h) {
            return (_$_669e[3] + h).substr(0, 1) === _$_669e[18]
        },
        ra = function(h) {
            var c = this.options.columns;
            return c[h] && c[h].name ? c[h].name : h
        },
        R = function(h, c, b) {
            var a;
            if (a = this.options.data[c]) {
                var d =
                    0 <= h ? ra.call(this, h) : h;
                typeof b === _$_669e[21] ? b = typeof d === _$_669e[22] ? a[h] : jSuites.path.call(a, _$_669e[3] + d) : (typeof d === _$_669e[22] ? a[h] = b : jSuites.path.call(a, _$_669e[3] + d, b), !Z(b) && this.records[c][h] && (this.records[c][h].v = b))
            } else b = null;
            return b
        },
        T = function(h) {
            return 0 == this.config.editable || h && 0 == h.options.editable ? !1 : (h = Oa.call(this)) ? 7 <= h ? !0 : !1 : !0
        },
        Pa = function(h) {
            var c = _$_669e[23];
            if (this.parent) var b = this.parent.config,
                a = this.parent.element,
                d = this;
            else b = this.config, a = this.element, d = r.current ?
                r.current : null;
            typeof h === _$_669e[21] && (h = !a.classList.contains(c));
            h ? (a.classList.add(c), b.fullscreen = !0) : (a.classList.remove(c), b.fullscreen = !1);
            d && (J.call(d, !0), ma.update.call(d.parent, d))
        },
        Qa = function(h) {
            var c = this.loading;
            null === h && (h = c.style.display == _$_669e[24] ? !0 : !1);
            c.style.display = h ? _$_669e[24] : _$_669e[3];
            setTimeout(function() {
                c.style.display = _$_669e[3]
            }, 1E3)
        },
        Ra = function(h, c) {
            if (typeof h === _$_669e[25]) {
                var b = H.getCoordsFromColumnName(h);
                b = this.records[b[1]][b[0]];
                if (b.element) h = b.element;
                else return typeof c === _$_669e[21] ? b.readonly : b.readonly = c ? !0 : !1
            }
            b = _$_669e[26];
            h = h.classList;
            if (typeof c === _$_669e[21]) return h.contains(b) ? !0 : !1;
            c ? h.add(b) : h.remove(b)
        },
        Ya = function(h, c) {
            var b, a = [];
            if (typeof h === _$_669e[25]) {
                var d = H.getCoordsFromColumnName(h);
                a.push({
                    x: d[0],
                    y: d[1],
                    value: {
                        rotate: c
                    }
                })
            } else if (Array.isArray(h))
                for (b = 0; b < h.length; b++) d = H.getCoordsFromColumnName(h[b]), a.push({
                    x: d[0],
                    y: d[1],
                    value: {
                        rotate: c
                    }
                });
            for (b = 0; b < a.length; b++)(d = this.records[a[b].y][a[b].x]) && d.element && (h = d.element.firstChild,
                c ? (h && h.classList && h.classList.contains(_$_669e[27]) || (h = document.createElement(_$_669e[28]), h.classList.add(_$_669e[27]), h.innerHTML = d.element.innerHTML, d.element.innerHTML = _$_669e[3], d.element.appendChild(h)), h.style.transform = _$_669e[29] + parseInt(270 - c) + _$_669e[30]) : h && h.classList && h.classList.contains(_$_669e[27]) && (h.classList.remove(_$_669e[27]), h.style.transform = _$_669e[3]));
            this.updateProperty(a);
            ca.refresh.call(this)
        },
        Ca = function(h) {
            var c = _$_669e[31],
                b = this.table.classList;
            if (typeof h == _$_669e[21]) return b.contains(c) ?
                !1 : !0;
            h ? b.remove(c) : b.add(c);
            J.call(this, !0)
        },
        Oa = function(h) {
            if (void 0 === h) return this.status;
            var c = document.createElement(_$_669e[32]);
            c.style.textAlign = _$_669e[33];
            c.style.fontSize = _$_669e[34];
            c.style.cursor = _$_669e[35];
            c.onclick = function() {
                window.location.href = M().host + _$_669e[36]
            };
            var b = null,
                a = [],
                d = [74, 50, 48, 50, 48, 33].join(_$_669e[3]);
            (function(g) {
                b = 1;
                try {
                    if (g) {
                        var f = window.atob(g);
                        f = f.split(_$_669e[37]);
                        if (f[1]) {
                            f[1] = window.atob(f[1]);
                            var k = f[0];
                            g = d;
                            var l = f[1],
                                n = jSuites.sha512,
                                p = _$_669e[3],
                                q =
                                _$_669e[3];
                            128 < g.length && (g = n(g));
                            for (var v = 0; 128 > v; v++) {
                                var w = g[v] ? g[v].charCodeAt(0) : 0;
                                p += String.fromCharCode(54 ^ w);
                                q += String.fromCharCode(92 ^ w)
                            }
                            var x = n(q + n(p + l));
                            if (k != x) b = 3;
                            else if (f[1] = JSON.parse(f[1]), f[1].date) {
                                var y = window.location.hostname || _$_669e[38];
                                k = !1;
                                for (x = 0; x < f[1].domain.length; x++) - 1 === f[1].domain[x].indexOf(_$_669e[39]) ? y === f[1].domain[x] && (k = !0) : y.substr(y.indexOf(f[1].domain[x])) == f[1].domain[x] && (k = !0);
                                !1 === k ? b = 4 : f[1].scope && 0 <= f[1].scope.indexOf(_$_669e[40]) ? (y = new Date, k = parseInt(y.getTime() /
                                    1E3), x = M(), f[1].date < k ? (b = f[1].demo || !f[1].plan || f[1].date + 2592E3 < k ? 6 : 7, x.license = f[1].name + _$_669e[41]) : (b = 8, x.license = _$_669e[42] + f[1].name, x.edition = 3 == f[1].plan || 6 == f[1].plan || 31 == f[1].plan || 33 == f[1].plan ? _$_669e[43] : _$_669e[5], a = f[1].scope)) : b = 5
                            } else b = 4
                        } else b = 2
                    }
                } catch (B) {}
            })(h);
            h = _$_669e[3];
            if (1 == b) h = _$_669e[44];
            else if (2 == b || 3 == b || 4 == b) h = _$_669e[45];
            else if (5 == b) h = _$_669e[46];
            else if (6 == b || 7 == b) h = _$_669e[47];
            c.appendChild(document.createTextNode(h));
            if (8 > b) try {
                if (typeof sessionStorage !== _$_669e[21] &&
                    !sessionStorage.getItem(_$_669e[48])) {
                    sessionStorage.setItem(_$_669e[48], !0);
                    var e = document.createElement(_$_669e[49]);
                    e.src = M().host + _$_669e[50] + _$_669e[51];
                    e.style.display = _$_669e[52];
                    c.appendChild(e)
                }
            } catch (g) {}
            Object.defineProperty(this, _$_669e[53], {
                value: b,
                writable: !1,
                configurable: !1,
                enumerable: !1,
                extensible: !1
            });
            Object.defineProperty(this, _$_669e[54], {
                value: M().edition === _$_669e[43] ? 1 : 0,
                writable: !1,
                configurable: !1,
                enumerable: !1,
                extensible: !1
            });
            Object.defineProperty(this, _$_669e[55], {
                get: function() {
                    return a
                },
                configurable: !1,
                enumerable: !1,
                extensible: !1
            });
            return c
        },
        J = function() {
            var h = function(l) {
                    -1 == this.indexOf(l) && this.push(l)
                },
                c = function(l, n) {
                    if (!l || !n) return !1;
                    var p = l.length;
                    if (p != n.length) return !1;
                    for (; p--;)
                        if (l[p] !== n[p]) return !1;
                    return !0
                },
                b = function(l) {
                    var n = null,
                        p, q, v = 0,
                        w = {};
                    for (p = 0; p < this.options.columns.length; p++)
                        if (-1 === l.indexOf(p)) {
                            if (na.attached.call(this, p) && (this.headers[p].parentNode.removeChild(this.headers[p]), this.colgroup[p].parentNode.removeChild(this.colgroup[p]), this.options.footers))
                                for (q =
                                    0; q < this.options.footers.length; q++) this.footers.content[q][p].element && this.footers.content[q][p].element.parentNode && this.footers.content[q][p].element.parentNode.removeChild(this.footers.content[q][p].element);
                            for (q = 0; q < this.visibleRows.length; q++) {
                                var x = this.records[this.visibleRows[q]][p];
                                x.element && x.element.remove()
                            }
                        } if (this.options.nestedHeaders)
                        for (q = 0; q < this.thead.children.length - 1; q++) v += this.thead.children[q].offsetHeight;
                    for (p = 0; p < l.length; p++) {
                        var y = l[p];
                        var B = Fa.create.call(this, y);
                        if (n) {
                            if (!B.parentNode) {
                                this.headerContainer.insertBefore(B, this.headerContainer.children[n].nextSibling);
                                this.colgroupContainer.insertBefore(this.colgroup[y], this.colgroupContainer.children[n].nextSibling);
                                if (this.options.footers)
                                    for (q = 0; q < this.options.footers.length; q++) this.tfoot.children[q].insertBefore(ia.create.call(this, y, q), this.tfoot.children[q].children[n].nextSibling);
                                v && (B.style.top += v)
                            }
                            for (q = 0; q < this.visibleRows.length; q++) x = this.visibleRows[q], x = U[_$_669e[56]].call(this, y, x), x.parentNode ||
                                this.tbody.children[q].insertBefore(x, this.tbody.children[q].children[n].nextSibling)
                        } else {
                            if (!B.parentNode) {
                                this.headerContainer.insertBefore(B, this.headerContainer.children[1]);
                                this.colgroupContainer.insertBefore(this.colgroup[y], this.colgroupContainer.children[1]);
                                if (this.options.footers)
                                    for (q = 0; q < this.options.footers.length; q++) this.tfoot.children[q].insertBefore(ia.create.call(this, y, q), this.tfoot.children[q].children[1]);
                                v && (B.style.top += v)
                            }
                            for (q = 0; q < this.visibleRows.length; q++) x = this.visibleRows[q],
                                x = U[_$_669e[56]].call(this, y, x), x.parentNode || this.tbody.children[q].insertBefore(x, this.tbody.children[q].children[1])
                        }
                        if (this.options.mergeCells)
                            for (q = 0; q < this.visibleRows.length; q++) x = this.visibleRows[q], n = H.getColumnNameFromCoords(y, x), this.options.mergeCells[n] && (this.records[x][y].merged || (w[n] = !1));
                        n = Array.prototype.indexOf.call(this.headerContainer.children, B)
                    }
                    v = 0;
                    if (this.options.nestedHeaders)
                        for (q = 0; q < this.options.nestedHeaders.length; q++) {
                            n = null;
                            y = [];
                            x = Aa.getColumns.call(this, q);
                            for (p =
                                0; p < l.length; p++) void 0 !== x[l[p]] && -1 == y.indexOf(x[l[p]]) && y.push(x[l[p]]);
                            for (p = 0; p < this.options.nestedHeaders[q].length; p++) - 1 == y.indexOf(p) && this.nested.content[q][p] && this.nested.content[q][p].element.offsetHeight && this.nested.content[q][p].element.remove();
                            for (p = x = 0; p < y.length; p++) B = Aa.create.call(this, y[p], q), B.parentNode || (n ? this.thead.children[q].children[n] ? this.thead.children[q].insertBefore(B, this.thead.children[q].children[n].nextSibling) : this.thead.children[q].appendChild(B) : this.thead.children[q].insertBefore(B,
                                this.thead.children[q].children[1]), B.style.top = v + _$_669e[57], this.options.freezeColumns && this.options.nestedHeaders[q][y[p]].frozen && (B.style.left = x + (Ca.call(this) ? 50 : 0) + _$_669e[57], B.classList.add(_$_669e[58]))), x += B.offsetWidth, n = Array.prototype.indexOf.call(this.thead.children[q].children, B);
                            v += this.thead.children[q].offsetHeight
                        }
                    sa.batch.call(this, w);
                    this.visibleCols = l
                },
                a = function(l) {
                    var n = [],
                        p = null,
                        q;
                    for (q = 0; q < this.tbody.children.length; q++) n.push(this.tbody.children[q]);
                    for (q = 0; q < n.length; q++) {
                        var v =
                            parseInt(n[q].getAttribute(_$_669e[59])); - 1 === l.indexOf(v) && n[q].remove()
                    }
                    for (q = 0; q < l.length; q++) n = aa.create.call(this, l[q]), n.parentNode || (p ? this.tbody.insertBefore(n, p.nextSibling) : this.tbody.insertBefore(n, this.tbody.firstChild)), p = n;
                    this.visibleRows = l
                },
                d = function() {
                    var l = [],
                        n = 0,
                        p = null,
                        q = this.content.scrollLeft,
                        v = this.content.scrollWidth,
                        w = this.content.offsetWidth,
                        x = this.options.freezeColumns || 0,
                        y = J.limited.call(this) && 0 != this.options.virtualizationX;
                    if (y) {
                        var B = p = 0;
                        p = this.options.columns;
                        B = q;
                        var A = B + w,
                            F = null,
                            u = 0;
                        x && (B += ea.getWidth.call(this, !0));
                        for (n = 0; n < p.length; n++) F = na.isVisible.call(this, n) ? parseInt(p[n].width) : 0, (u + F > B && u <= A || n < x) && h.call(l, n), u += F;
                        for (n = 0; n < l.length; n++)
                            if (B = sa.getAffected.call(this.merged.cols, l[n]))
                                for (p = 0; p < B.length; p++) h.call(l, B[p]);
                        for (n = 0; n < l.length; n++)
                            if (B = Aa.getAffected.call(this, l[n]))
                                for (p = 0; p < B.length; p++) this.options.columns[B[p]] && h.call(l, B[p]);
                        l.sort(function(z, C) {
                            return z - C
                        })
                    } else
                        for (p = this.options.columns, n = 0; n < p.length; n++) l.push(n);
                    c(l, this.visibleCols) ||
                        b.call(this, l);
                    y && (this.table.style.left = q + _$_669e[57], l = g.width.call(this, l[x]) - q, -100 < l && v - (q + w) > w / 2 && (l = 0), this.table.style.marginLeft = l + _$_669e[57])
                },
                e = function() {
                    var l = [],
                        n = null,
                        p = 0,
                        q = 0;
                    p = null;
                    var v = 0,
                        w = this.options.freezeRows || 0,
                        x = this.content.scrollTop,
                        y = this.content.scrollHeight,
                        B = this.content.offsetHeight,
                        A = J.limited.call(this) && !this.options.pagination && 0 != this.options.virtualizationY;
                    if (A) {
                        var F = x,
                            u = F + this.content.offsetHeight,
                            z = 0,
                            C = 0;
                        n = this.results ? this.results : this.rows;
                        w && (F += ja.getHeight.call(this,
                            !0));
                        for (q = 0; q < n.length; q++) p = this.results ? n[q] : q, !1 !== this.rows[p].visible && (z = aa.height.call(this, p), (C + z > F && C < u || q < w) && l.push(p), C += z)
                    } else
                        for (n = this.results ? this.results : this.rows, this.options.pagination ? (p = parseInt(this.options.pagination) * this.pageNumber, v = p + parseInt(this.options.pagination), v > n.length && (v = n.length)) : (p = 0, v = n.length), q = p; q < v; q++) p = this.results ? n[q] : q, l.push(p);
                    if (this.options.mergeCells) {
                        for (q = 0; q < l.length; q++)
                            if (this.merged.rows[l[q]] && (v = sa.getAffected.call(this.merged.rows,
                                    l[q])))
                                for (p = 0; p < v.length; p++) - 1 === l.indexOf(v[p]) && l.push(v[p]);
                        l.sort(function(E, K) {
                            return E - K
                        })
                    }
                    c(l, this.visibleRows) || (a.call(this, l), this.visibleCols = null);
                    A && (this.table.style.top = x + _$_669e[57], l = g.height.call(this, l[w]) - x, -24 < l && y - (x + B) > B / 2 && (l = 0), this.table.style.marginTop = l + _$_669e[57])
                },
                g = function(l, n) {
                    n && (this.visibleCols = this.visibleRows = null);
                    g.calculate.call(this);
                    g.update.call(this);
                    if (!0 !== l) {
                        var p = this;
                        this.content.addEventListener(_$_669e[60], function() {
                            g.ignore || (g.update.call(p),
                                ca.refresh.call(p))
                        });
                        this.content.addEventListener(_$_669e[61], function(q) {
                            if (J.limited.call(p) && !p.options.pagination) {
                                var v = Math.abs(q.deltaX),
                                    w = Math.abs(q.deltaY);
                                q.shiftKey || v > w ? 0 > q.deltaX || 0 > q.deltaY ? g.left.call(p) : g.right.call(p) : 0 > q.deltaY ? g.up.call(p) : g.down.call(p);
                                q.preventDefault()
                            }
                        })
                    }
                };
            g.calculate = function() {
                if (J.limited.call(this)) {
                    var l = ea.getWidth.call(this),
                        n = ja.getHeight.call(this);
                    this.width = g.width.call(this) + (l ? l : 1);
                    this.height = g.height.call(this) + (n ? n : 1);
                    l = this.parent.config.spacing ||
                        100;
                    this.scroll.style.width = this.width + l + _$_669e[57];
                    this.scroll.style.height = this.height + l + _$_669e[57]
                }
            };
            g.update = function() {
                var l = this.thead.lastChild.offsetHeight;
                e.call(this);
                d.call(this);
                if (J.limited.call(this)) {
                    var n = null;
                    var p = this.thead.lastChild.offsetHeight;
                    l < p && (this.thead.lastChild.style.height = p + _$_669e[57], n = !0);
                    for (l = 0; l < this.visibleRows.length; l++) {
                        var q = this.rows[this.visibleRows[l]];
                        p = q.element.offsetHeight;
                        if (p > q.height || !q.height) q.element.style.height = p + _$_669e[57], q.height = p,
                            n = !0
                    }
                    n && (p = this.parent.config.spacing || 100, n = ja.getHeight.call(this), this.height = g.height.call(this) + (n ? n : 1), this.scroll.style.height = this.height + p + _$_669e[57])
                }
                Ea.update.call(this);
                ea.refresh.call(this);
                ja.refresh.call(this)
            };
            g[_$_669e[62]] = function(l, n) {
                null !== l && (l = g.height.call(this, l), this.content.scrollTop = l);
                void 0 !== n && null !== n && (n = g.width.call(this, n), this.content.scrollLeft = n)
            };
            g.up = function(l) {
                var n = this.options.freezeRows || 0;
                if (l) this.content.scrollTop = 0;
                else if (this.tbody.children[n +
                        1]) {
                    l = this.tbody.children[n + 1].getAttribute(_$_669e[59]);
                    l = parseInt(l);
                    var p = l;
                    for (this.results && (l = this.results.indexOf(l)); 0 < l;) {
                        l--;
                        var q = this.results ? this.results[l] : l;
                        if (aa.isVisible.call(this, q)) {
                            p = q;
                            break
                        }
                    }
                    q = aa.height.call(this, p);
                    l = this.tbody.children[n].getAttribute(_$_669e[59]);
                    aa.attached.call(this, l) || (q += this.content.scrollTop - this.table.offsetTop);
                    this.content.scrollTop -= q
                }
            };
            g.down = function(l) {
                var n = this.options.freezeRows || 0;
                l ? this.content.scrollTop = this.content.scrollHeight : this.tbody.firstChild &&
                    (this.content.scrollTop += this.tbody.children[1 + n].offsetHeight)
            };
            g.left = function(l) {
                var n = this.options.freezeColumns || 0;
                if (l) this.content.scrollLeft = 0;
                else if (n = this.thead.lastChild.children[1 + n]) {
                    for (n = l = this.thead.lastChild.children[1 + n].getAttribute(_$_669e[63]); 0 < l;)
                        if (l--, na.isVisible.call(this, l)) {
                            n = l;
                            break
                        } l = n;
                    n = parseInt(this.options.columns[l].width);
                    this.headers[l] && this.headers[l].parentNode || (n += this.content.scrollLeft - this.table.offsetLeft);
                    this.content.scrollLeft -= n
                }
            };
            g.right = function(l) {
                var n =
                    this.options.freezeColumns || 0;
                if (l) this.content.scrollLeft = this.content.scrollWidth;
                else if (n = this.thead.lastChild.children[1 + n]) this.content.scrollLeft += n.offsetWidth
            };
            var f = function() {
                    for (var l, n = this.content.offsetHeight - this.thead.offsetHeight - ja.getHeight.call(this), p = 0, q = 0; q < this.tbody.children.length; q++) l = this.tbody.children[q].offsetHeight, n > l && (p += l, n -= l);
                    return p
                },
                k = function() {
                    for (var l, n = this.content.offsetWidth - ea.getWidth.call(this), p = 0, q = 0; q < this.thead.lastChild.children.length; q++) l =
                        this.thead.lastChild.children[q].offsetWidth, n > l && (p += l, n -= l);
                    return p
                };
            g.pageUp = function() {
                this.content.scrollTop -= f.call(this)
            };
            g.pageDown = function() {
                this.content.scrollTop += f.call(this)
            };
            g.pageLeft = function() {
                this.content.scrollLeft -= k.call(this)
            };
            g.pageRight = function() {
                this.content.scrollLeft += k.call(this)
            };
            g.build = function() {
                var l = document.createElement(_$_669e[32]);
                l.className = _$_669e[64];
                return l
            };
            g.width = function(l, n) {
                for (var p = this.options.columns, q = 0, v = this.options.freezeColumns || 0; v < this.options.columns.length; v++) {
                    if (l ===
                        v) return q;
                    na.isVisible.call(this, v) && (q += parseInt(p[v].width));
                    if (q >= n) return v
                }
                return n ? v : q
            };
            g.height = function(l, n) {
                for (var p = 0, q = 0, v = this.results ? this.results : this.rows, w = this.options.freezeRows || 0; w < v.length; w++) {
                    q = w;
                    this.results && (q = v[w]);
                    if (l === q) return p;
                    this.rows[q] && !1 !== this.rows[q].visible && (p += aa.height.call(this, q));
                    if (p >= n) return q
                }
                return n ? q : p
            };
            g.adjust = function(l, n, p) {
                var q = this.records[n][l].element;
                if (q) {
                    if (!this.records[n][l].merged)
                        if (0 == p || 2 == p) {
                            var v = 0,
                                w = Ca.call(this) ? 50 : 0;
                            this.options.freezeColumns &&
                                (v = ea.getWidth.call(this, !0));
                            n = this.content.scrollLeft + v;
                            v = this.content.scrollLeft + (this.content.offsetWidth - w);
                            w = this.table.offsetLeft + q.offsetLeft - w;
                            var x = q.offsetWidth;
                            0 == p ? w + x < v && l > this.options.freezeColumns - 1 && w < n && (q.offsetHeight ? (p = this.content.scrollLeft + (w - n), g.setX.call(this, p)) : g[_$_669e[62]].call(this, null, l)) : l > this.options.freezeColumns - 1 && (w > n ? w + x > v && (q.parentNode ? (p = this.content.scrollLeft + (w + x - v + 16), g.setX.call(this, p)) : g[_$_669e[62]].call(this, null, l)) : g[_$_669e[62]].call(this,
                                null, l))
                        } else v = 0, this.options.freezeRows && (v = ja.getHeight.call(this, !0)), l = this.content.scrollTop + v, v = this.content.scrollTop + this.content.offsetHeight - this.thead.offsetHeight - 16, w = this.table.offsetTop + q.offsetTop - this.thead.offsetHeight, x = q.offsetHeight, 1 == p ? w + x < v && n > this.options.freezeRows - 1 && w < l && (q.offsetHeight ? (p = this.content.scrollTop + (w - l), g.setY.call(this, p)) : g[_$_669e[62]].call(this, n)) : n > this.options.freezeRows - 1 && (w > l ? w + x > v && (q.parentNode ? (p = this.content.scrollTop + (w + x - v), g.setY.call(this,
                            p)) : g[_$_669e[62]].call(this, n)) : g[_$_669e[62]].call(this, n, null))
                } else 0 == p || 2 == p ? g[_$_669e[62]].call(this, null, l) : g[_$_669e[62]].call(this, n)
            };
            g.setX = function(l) {
                g.ignore = !0;
                this.content.scrollLeft = l;
                g.ignore = !1
            };
            g.setY = function(l) {
                g.ignore = !0;
                this.content.scrollTop = l;
                g.ignore = !1
            };
            g.limited = function() {
                return 1 == this.options.tableOverflow || 1 == this.parent.config.fullscreen
            };
            return g
        }(),
        Sa = function(h, c, b, a) {
            var d = c,
                e = null;
            if (this.results && (c = this.results.indexOf(c), -1 == c)) return d;
            for (; 0 < c;) {
                c--;
                var g =
                    this.results ? this.results[c] : c;
                if (!1 !== this.rows[g].visible) {
                    var f = this.records[g][h].v === _$_669e[3] ? !1 : !0;
                    null === e && (e = f);
                    if (this.records[g][h].element && this.records[g][h].element.classList.contains(_$_669e[65])) {
                        if (this.records[g][h].merged) {
                            d = g - this.records[g][h].merged[1];
                            if (!b) break;
                            f = !0
                        }
                    } else {
                        if (b && f !== e && !a) break;
                        if (b) d = this.records[g][h].merged ? g - this.records[g][h].merged[1] : g;
                        else return g
                    }
                    e = f
                }
            }
            return d
        },
        Ta = function(h, c, b, a) {
            var d = c,
                e = null;
            if (this.results) {
                c = this.results.indexOf(c);
                if (-1 ==
                    c) return d;
                var g = this.results.length - 1
            } else g = this.rows.length - 1;
            for (; c < g;) {
                c++;
                var f = this.results ? this.results[c] : c;
                if (!1 !== this.rows[f].visible) {
                    var k = this.records[f][h].v === _$_669e[3] ? !1 : !0;
                    null === e && (e = k);
                    if (this.records[f][h].element && this.records[f][h].element.classList.contains(_$_669e[65])) {
                        if (this.records[f][h].merged) {
                            if (!this.records[f][h].merged[1] && (d = f, !b)) break;
                            k = !0
                        }
                    } else {
                        if (b && k !== e && !a) break;
                        if (b) d = f;
                        else return f
                    }
                    e = k
                }
            }
            return d
        },
        Ua = function(h, c, b, a) {
            for (var d = this.options.columns,
                    e = h, g, f = null; h < d.length - 1;)
                if (h++, na.isVisible.call(this, h)) {
                    g = !1;
                    this.records[c] && (g = this.records[c][h].merged ? this.records[c - this.records[c][h].merged[1]][h - this.records[c][h].merged[0]].v ? !0 : !1 : this.records[c][h].v ? !0 : !1);
                    null === f && (f = g);
                    if (this.records[c] && this.records[c][h].element && this.records[c][h].element.classList.contains(_$_669e[65])) {
                        if (this.records[c][h].merged && h == h + this.records[c][h].merged[0] && (e = h, !b)) break
                    } else {
                        if (b && g !== f && !a) break;
                        if (b) e = h;
                        else return h
                    }
                    f = g
                } return e
        },
        Va = function(h,
            c, b, a) {
            for (var d = h, e, g = null; 0 < h;)
                if (h--, na.isVisible.call(this, h)) {
                    e = !1;
                    this.records[c] && (e = this.records[c][h].merged ? this.records[c - this.records[c][h].merged[1]][h - this.records[c][h].merged[0]].v ? !0 : !1 : this.records[c][h].v ? !0 : !1);
                    null === g && (g = e);
                    if (this.records[c] && this.records[c][h].element && this.records[c][h].element.classList.contains(_$_669e[65])) {
                        if (this.records[c][h].merged && (d = h - this.records[c][h].merged[0], !b)) break
                    } else {
                        if (b && e !== g && !a) break;
                        if (b) d = h;
                        else return h
                    }
                    g = e
                } return d
        },
        Za = function() {
            var h =
                function(b, a, d, e) {
                    1 == b || 3 == b ? aa.attached.call(this, d) || (0 < this.options.pagination ? this.page(this.whichPage(d)) : 1 == b ? J.up.call(this, e) : J.down.call(this, e)) : U.attached.call(this, a, d) || (0 == b ? J.left.call(this, e) : J.right.call(this, e))
                },
                c = function(b) {
                    b.up = c.up;
                    b.down = c.down;
                    b.right = c.right;
                    b.left = c.left;
                    b.last = c.last;
                    b.first = c.first
                };
            c.up = function(b, a, d) {
                if (Array.isArray(this.selectedCell)) {
                    if (b) var e = parseInt(this.selectedCell[2]),
                        g = parseInt(this.selectedCell[3]);
                    else e = parseInt(this.selectedCell[0]),
                        g = parseInt(this.selectedCell[1]);
                    d = Sa.call(this, e, g, a, d);
                    d != g && (g = d, b ? (b = this.selectedCell[0], d = this.selectedCell[1]) : (b = e, d = g, e = b, g = d), h.call(this, 1, e, g, a), ka[_$_669e[11]].call(this, b, d, e, g), J.adjust.call(this, e, g, 1))
                }
            };
            c.down = function(b, a, d) {
                if (Array.isArray(this.selectedCell)) {
                    if (b) var e = parseInt(this.selectedCell[2]),
                        g = parseInt(this.selectedCell[3]);
                    else e = parseInt(this.selectedCell[0]), g = parseInt(this.selectedCell[1]);
                    d = Ta.call(this, e, g, a, d);
                    d != g && (g = d, b ? (b = this.selectedCell[0], d = this.selectedCell[1]) :
                        (b = e, d = g, e = b, g = d), h.call(this, 3, e, g, a), ka[_$_669e[11]].call(this, b, d, e, g), J.adjust.call(this, e, g, 3))
                }
            };
            c.right = function(b, a, d) {
                if (Array.isArray(this.selectedCell)) {
                    if (b) var e = parseInt(this.selectedCell[2]),
                        g = parseInt(this.selectedCell[3]);
                    else e = parseInt(this.selectedCell[0]), g = parseInt(this.selectedCell[1]);
                    d = Ua.call(this, e, g, a, d);
                    d != e && (e = d, b ? (b = this.selectedCell[0], d = this.selectedCell[1]) : (b = e, d = g, e = b, g = d), h.call(this, 2, e, g, a), ka[_$_669e[11]].call(this, b, d, e, g), J.adjust.call(this, e, g, 2))
                }
            };
            c.left =
                function(b, a, d) {
                    if (Array.isArray(this.selectedCell)) {
                        if (b) var e = parseInt(this.selectedCell[2]),
                            g = parseInt(this.selectedCell[3]);
                        else e = parseInt(this.selectedCell[0]), g = parseInt(this.selectedCell[1]);
                        d = Va.call(this, e, g, a, d);
                        d != e && (e = d, b ? (b = this.selectedCell[0], d = this.selectedCell[1]) : (b = e, d = g, e = b, g = d), h.call(this, 0, e, g, a), ka[_$_669e[11]].call(this, b, d, e, g), J.adjust.call(this, e, g, 0))
                    }
                };
            c.first = function(b, a) {
                Array.isArray(this.selectedCell) && (a ? c.up.call(this, b, !0, !0) : c.left.call(this, b, !0, !0))
            };
            c.last =
                function(b, a) {
                    Array.isArray(this.selectedCell) && (a ? c.down.call(this, b, !0, !0) : c.right.call(this, b, !0, !0))
                };
            return c
        }(),
        $a = function() {
            return function(h) {
                h.save = Wa;
                h.fullscreen = Pa;
                h.setReadOnly = Ra;
                h.isReadOnly = Ra;
                h.rotate = Ya;
                h.setPlugins = function(c) {
                    wa.call(this.parent, c)
                };
                h.showIndex = function() {
                    Ca.call(this, !0)
                };
                h.hideIndex = function() {
                    Ca.call(this, !1)
                };
                h[_$_669e[62]] = J[_$_669e[62]];
                h.undo = function() {
                    Q.undo.call(h.parent)
                };
                h.redo = function() {
                    Q.redo.call(h.parent)
                };
                h.isEditable = function() {
                    return T.call(h.parent,
                        h)
                };
                h.dispatch = function() {
                    return G.apply(this.parent, arguments)
                };
                h.persistence = L;
                h.setViewport = function(c, b) {
                    this.options.tableOverflow = !0;
                    c = parseInt(c);
                    b = parseInt(b);
                    this.content.classList.add(_$_669e[66]);
                    150 < c && (this.options.tableWidth = c + _$_669e[57], this.content.style.width = this.options.tableWidth);
                    150 < b && !this.options.pagination && (this.options.tableHeight = parseInt(b) + _$_669e[57], this.content.style.height = this.options.tableHeight);
                    J.call(this, !0);
                    ma.update.call(this.parent, this);
                    G.call(this.parent,
                        _$_669e[67], c, b)
                };
                h.helpers = H
            }
        }(),
        Ha = function() {
            var h = {
                    logo: null,
                    url: null,
                    persistence: !1,
                    sequence: !0,
                    data: null,
                    json: null,
                    rows: [],
                    columns: [],
                    cells: {},
                    role: null,
                    nestedHeaders: null,
                    defaultColWidth: 100,
                    defaultRowHeight: 0,
                    defaultColAlign: null,
                    minSpareRows: 0,
                    minSpareCols: 0,
                    minDimensions: [0, 0],
                    csv: null,
                    csvFileName: _$_669e[48],
                    csvHeaders: !0,
                    csvDelimiter: _$_669e[37],
                    columnSorting: !0,
                    columnSortingOnDblClick: !0,
                    columnDrag: !0,
                    columnResize: !0,
                    rowResize: !0,
                    rowDrag: !0,
                    editable: !0,
                    allowInsertRow: !0,
                    allowManualInsertRow: !0,
                    allowInsertColumn: !0,
                    allowManualInsertColumn: !0,
                    allowDeleteRow: !0,
                    allowDeletingAllRows: !1,
                    allowDeleteColumn: !0,
                    allowRenameColumn: !0,
                    allowComments: !0,
                    selectionCopy: !0,
                    mergeCells: {},
                    search: !1,
                    pagination: !1,
                    paginationOptions: null,
                    textOverflow: !1,
                    tableOverflow: !1,
                    virtualizationY: !0,
                    virtualizationX: !0,
                    tableHeight: null,
                    tableWidth: null,
                    resize: null,
                    comments: null,
                    meta: null,
                    style: {},
                    freezeColumns: 0,
                    freezeRows: 0,
                    freezeColumnControl: !1,
                    freezeRowControl: !1,
                    orderBy: null,
                    worksheetId: _$_669e[3],
                    worksheetName: null,
                    worksheetState: null,
                    filters: !1,
                    footers: null,
                    formify: null,
                    applyMaskOnFooters: !1,
                    pluginOptions: null,
                    locked: !1,
                    selectUnLockedCells: !0,
                    selectLockedCells: !0,
                    wrap: !1,
                    gridline: !0
                },
                c = {
                    application: _$_669e[68],
                    cloud: null,
                    root: null,
                    definedNames: {},
                    sorting: null,
                    server: null,
                    toolbar: null,
                    editable: !0,
                    allowExport: !0,
                    includeHeadersOnDownload: !1,
                    forceUpdateOnPaste: !1,
                    loadingSpin: !1,
                    fullscreen: !1,
                    secureFormulas: !0,
                    parseFormulas: !0,
                    debugFormulas: !1,
                    editorFormulas: !0,
                    autoIncrement: !0,
                    autoCasting: !0,
                    stripHTML: !1,
                    tabs: !1,
                    allowDeleteWorksheet: !0,
                    allowRenameWorksheet: !0,
                    allowMoveWorksheet: !0,
                    moveDownOnEnter: !0,
                    spacing: 100,
                    onevent: null,
                    onclick: null,
                    onload: null,
                    onundo: null,
                    onredo: null,
                    onbeforesave: null,
                    onsave: null,
                    onbeforechange: null,
                    onchange: null,
                    onafterchanges: null,
                    oncopy: null,
                    onbeforepaste: null,
                    onpaste: null,
                    onbeforeinsertrow: null,
                    oninsertrow: null,
                    onbeforedeleterow: null,
                    ondeleterow: null,
                    onbeforeinsertcolumn: null,
                    oninsertcolumn: null,
                    onbeforedeletecolumn: null,
                    ondeletecolumn: null,
                    onmoverow: null,
                    onmovecolumn: null,
                    onresize: null,
                    onresizerow: null,
                    onresizecolumn: null,
                    onselection: null,
                    onbeforecomments: null,
                    oncomments: null,
                    onbeforesort: null,
                    onsort: null,
                    onfocus: null,
                    onblur: null,
                    onmerge: null,
                    onchangeheader: null,
                    onchangefooter: null,
                    onchangefootervalue: null,
                    onrenderfootercell: null,
                    onchangenested: null,
                    onchangenestedcell: null,
                    oncreateeditor: null,
                    oneditionstart: null,
                    oneditionend: null,
                    onchangestyle: null,
                    onchangemeta: null,
                    onbeforechangepage: null,
                    onchangepage: null,
                    onbeforecreateworksheet: null,
                    oncreateworksheet: null,
                    onrenameworksheet: null,
                    ondeleteworksheet: null,
                    onmoveworksheet: null,
                    onopenworksheet: null,
                    onchangerowid: null,
                    onbeforesearch: null,
                    onsearchstart: null,
                    onsearch: null,
                    onsearchrow: null,
                    onbeforefilter: null,
                    onfilter: null,
                    onopenfilter: null,
                    oncreatecell: null,
                    oncreaterow: null,
                    oncreatecolumn: null,
                    onbeforeformula: null,
                    onformulachain: null,
                    onchangereferences: null,
                    onbeforechangereferences: null,
                    onbeforesend: null,
                    onchangedefinednames: null,
                    oninput: null,
                    onchangerowvisibility: null,
                    onchangecolumnvisibility: null,
                    oncreaterowgroup: null,
                    onopenrowgroup: null,
                    oncloserowgroup: null,
                    ondestroyrowgroup: null,
                    updateTable: null,
                    contextMenu: null,
                    parseTableFirstRowAsHeader: !1,
                    parseTableAutoCellType: !1,
                    plugins: null,
                    about: !0,
                    license: null,
                    worksheets: null,
                    validations: null
                },
                b = function(e, g) {
                    g = JSON.parse(JSON.stringify(g));
                    var f = {},
                        k;
                    for (k in g) e && e.hasOwnProperty(k) ? f[k] = e[k] : f[k] = g[k];
                    return f
                },
                a = function(e) {
                    e && (typeof e == _$_669e[25] && (e = JSON.parse(e)), this.options.data = e);
                    this.options.json && (this.options.data = this.options.json);
                    this.options.data || (this.options.data = []);
                    e = this.options.columns.length;
                    if (this.options.data && typeof this.options.data[0] !== _$_669e[21]) {
                        var g = this.options.data[0].data && 0 < Object.keys(this.options.data[0].data).length ? Object.keys(this.options.data[0].data) : Object.keys(this.options.data[0]);
                        g.length > e && (e = g.length)
                    }
                    this.options.minDimensions[0] > e && (e = this.options.minDimensions[0]);
                    e || (e = 8);
                    for (var f = 0; f < e; f++) this.options.columns[f] ? this.options.columns[f].type || (this.options.columns[f].type = _$_669e[69]) : this.options.columns[f] = {
                            type: _$_669e[69]
                        },
                        this.options.columns[f].type == _$_669e[70] && (this.options.columns[f].type = _$_669e[71], this.options.columns[f].autocomplete = !0), !this.options.columns[f].name && g && g[f] && Number(g[f]) != g[f] && (this.options.columns[f].name = g[f]), this.options.columns[f].width || (this.options.columns[f].width = parseInt(this.options.defaultColWidth) + _$_669e[57]), this.options.columns[f].type != _$_669e[71] || this.options.columns[f].source || (this.options.columns[f].source = []), this.options.columns[f].type == _$_669e[71] && this.options.columns[f].url &&
                        jSuites.ajax({
                            url: this.options.columns[f].url,
                            index: f,
                            method: _$_669e[72],
                            dataType: _$_669e[73],
                            worksheet: this,
                            group: this.parent.name,
                            success: function(k) {
                                this.worksheet.options.columns[this.index].source = k
                            }
                        })
                },
                d = function(e) {
                    e.getConfig = d[_$_669e[56]];
                    e.setConfig = d[_$_669e[11]]
                };
            d.spreadsheet = function(e) {
                this.config = b(e, c);
                1 != this.config.secureFormulas && console.log(_$_669e[74])
            };
            d.worksheet = function(e) {
                var g = this;
                g.options = b(e, h);
                g.options.csv ? jSuites.ajax({
                    url: g.options.csv,
                    method: _$_669e[72],
                    dataType: _$_669e[69],
                    group: g.parent.name,
                    success: function(f) {
                        f = H.parseCSV(f, g.options.csvDelimiter);
                        if (1 == g.options.csvHeaders && 0 < f.length)
                            for (var k = f.shift(), l = 0; l < k.length; l++) g.options.columns[l] || (g.options.columns[l] = {
                                type: _$_669e[69],
                                width: g.options.defaultColWidth
                            }), typeof g.options.columns[l].title === _$_669e[21] && (g.options.columns[l].title = k[l]);
                        a.call(g, f)
                    }
                }) : g.options.url ? jSuites.ajax({
                    url: g.options.url,
                    method: _$_669e[72],
                    dataType: _$_669e[73],
                    group: g.parent.name,
                    success: function(f) {
                        a.call(g, f)
                    }
                }) : a.call(g)
            };
            d[_$_669e[56]] = function() {
                return this.options
            };
            d[_$_669e[11]] = function(e) {
                typeof e == _$_669e[25] && (e = JSON.parse(e));
                if (this.parent) {
                    g = Object.keys(e);
                    for (f = 0; f < g.length; f++) typeof h[g[f]] !== _$_669e[21] && (this.options[g[f]] = e[g[f]]);
                    this.content.style.resize = e.resize ? e.resize : _$_669e[3];
                    !0 === e.filters ? la.show.call(this) : !1 === e.filters && la.hide.call(this);
                    !0 === e.search ? xa.show.call(this) : !1 === e.search && xa.hide.call(this);
                    !0 === e.toolbar ? ma.show.call(this.parent) : !1 === e.toolbar && ma.hide.call(this.parent);
                    e.minDimensions && e.minDimensions[0] && (g = parseInt(e.minDimensions[0]) - this.headers.length, 0 < g && this.insertColumn(g), g = parseInt(e.minDimensions[1]) - this.rows.length, 0 < g && this.insertRow(g));
                    L.call(this, _$_669e[75], {
                        data: JSON.stringify(e)
                    })
                } else
                    for (var g = Object.keys(e), f = 0; f < g.length; f++) typeof c[g[f]] !== _$_669e[21] && (this.config[g[f]] = e[g[f]])
            };
            return d
        }(),
        Q = function() {
            var h = function(a) {
                    if (1 != this.ignoreHistory)
                        if (this.historyCascade) this.history[this.historyIndex].cascade || (this.history[this.historyIndex].cascade = []), this.history[this.historyIndex].cascade.push(a);
                        else {
                            var d = ++this.historyIndex;
                            this.history = this.history.slice(0, d + 1);
                            this.history[d] = a
                        }
                },
                c = function(a) {
                    var d = a.worksheet;
                    if (a.action == _$_669e[76]) d.deleteRow(a.insertBefore ? a.rowNumber : a.rowNumber + 1, a.numOfRows);
                    else if (a.action == _$_669e[77]) d.insertRow(a.numOfRows, a.rowNumber, 1, a.data), U.setAttributes.call(d, a.attributes);
                    else if (a.action == _$_669e[78]) d.deleteColumn(a.insertBefore ? a.columnNumber : a.columnNumber + 1, a.numOfColumns);
                    else if (a.action ==
                        _$_669e[79]) d.insertColumn(a.numOfColumns, a.columnNumber, 1, a.properties, a.data, a.extra), U.setAttributes.call(d, a.attributes);
                    else if (a.action == _$_669e[80]) d.moveRow(a.newValue, a.oldValue);
                    else if (a.action == _$_669e[81]) d.moveColumn(a.newValue, a.oldValue);
                    else if (a.action == _$_669e[82]) d.showRow(a.records);
                    else if (a.action == _$_669e[83]) d.hideRow(a.records);
                    else if (a.action == _$_669e[84]) d.showColumn(a.records);
                    else if (a.action == _$_669e[85]) d.hideColumn(a.records);
                    else if (a.action == _$_669e[86]) d.removeMerge(a.newValue,
                        !0), d.setMerge(a.oldValue);
                    else if (a.action == _$_669e[87]) d.setMerge(a.oldValue);
                    else if (a.action == _$_669e[88]) d.setStyle(a.oldValue, null, null, 1);
                    else if (a.action == _$_669e[89]) d.setStyle(a.oldValue);
                    else if (a.action == _$_669e[90]) d.setWidth(a.column, a.oldValue);
                    else if (a.action == _$_669e[91]) d.setHeight(a.row, a.oldValue);
                    else if (a.action == _$_669e[92]) d.setHeader(a.column, a.oldValue);
                    else if (a.action == _$_669e[93] || a.action == _$_669e[94] || a.action == _$_669e[95]) d[a.action](a.oldValue);
                    else if (a.action ==
                        _$_669e[96]) d.orderBy(a.column, a.direction ? 0 : 1, a.oldValue, a.previousState);
                    else if (a.action == _$_669e[97] || a.action == _$_669e[98] || a.action == _$_669e[99]) {
                        var e = [];
                        for (var g = 0; g < a.records.length; g++) {
                            var f = {
                                x: a.records[g].x,
                                value: a.records[g].oldValue
                            };
                            void 0 !== a.records[g].y && (f.y = a.records[g].y);
                            a.records[g].oldStyle && (f.style = a.records[g].oldStyle);
                            e.push(f)
                        }
                        a.action == _$_669e[97] ? d.setValue(e) : d.setProperty(e)
                    } else if (a.action == _$_669e[100]) d.renameWorksheet(a.index, a.oldValue);
                    else if (a.action ==
                        _$_669e[101]) d.moveWorksheet(a.t, a.f);
                    else if (a.action == _$_669e[102]) {
                        e = [];
                        for (g = 0; g < a.records.length; g++) e.push({
                            index: a.records[g].index,
                            value: a.records[g].oldValue
                        });
                        d.setDefinedNames(e)
                    } else a.action == _$_669e[103] ? d.resetRowGroup(a.row) : a.action == _$_669e[104] ? d.setRowGroup(a.row, a.elements) : a.action == _$_669e[105] ? d.closeRowGroup(a.row) : a.action == _$_669e[106] && d.openRowGroup(a.row);
                    G.call(this, _$_669e[107], d, a)
                };
            h.undo = function() {
                if (0 <= this.historyIndex) var a = this.history[this.historyIndex--];
                if (a) {
                    this.ignoreHistory = !0;
                    var d = a.worksheet;
                    d.openWorksheet();
                    c.call(this, a);
                    if (a.cascade && a.cascade.length)
                        for (var e = 0; e < a.cascade.length; e++) c.call(this, a.cascade[e]);
                    a.selection && d.updateSelectionFromCoords(a.selection);
                    this.ignoreHistory = !1
                }
            };
            var b = function(a) {
                var d = a.worksheet;
                if (a.action == _$_669e[76]) d.insertRow(a.numOfRows, a.rowNumber, a.insertBefore, a.data);
                else if (a.action == _$_669e[77]) d.deleteRow(a.rowNumber, a.numOfRows);
                else if (a.action == _$_669e[78]) d.insertColumn(a.numOfColumns, a.columnNumber,
                    a.insertBefore, a.properties, a.data);
                else if (a.action == _$_669e[79]) d.deleteColumn(a.columnNumber, a.numOfColumns);
                else if (a.action == _$_669e[80]) d.moveRow(a.oldValue, a.newValue);
                else if (a.action == _$_669e[81]) d.moveColumn(a.oldValue, a.newValue);
                else if (a.action == _$_669e[86]) d.setMerge(a.newValue);
                else if (a.action == _$_669e[87]) d.removeMerge(a.newValue);
                else if (a.action == _$_669e[88]) d.setStyle(a.newValue, null, null, 1);
                else if (a.action == _$_669e[89]) d.resetStyle(a.cells);
                else if (a.action == _$_669e[90]) d.setWidth(a.column,
                    a.newValue);
                else if (a.action == _$_669e[91]) d.setHeight(a.row, a.newValue);
                else if (a.action == _$_669e[92]) d.setHeader(a.column, a.newValue);
                else if (a.action == _$_669e[93] || a.action == _$_669e[94] || a.action == _$_669e[95]) d[a.action](a.newValue);
                else if (a.action == _$_669e[96]) d.orderBy(a.column, a.direction, a.newValue);
                else if (a.action == _$_669e[97] || a.action == _$_669e[98] || a.action == _$_669e[99] || a.action == _$_669e[82] || a.action == _$_669e[83] || a.action == _$_669e[84] || a.action == _$_669e[85]) d[a.action](a.records);
                else a.action ==
                    _$_669e[100] ? d.renameWorksheet(a.index, a.newValue) : a.action == _$_669e[101] ? d.moveWorksheet(a.f, a.t) : a.action == _$_669e[102] ? d.setDefinedNames(a.records) : a.action == _$_669e[103] ? d.setRowGroup(a.row, a.elements) : a.action == _$_669e[104] ? d.resetRowGroup(a.row) : a.action == _$_669e[105] ? d.openRowGroup(a.row) : a.action == _$_669e[106] && d.closeRowGroup(a.row);
                G.call(this, _$_669e[108], d, a)
            };
            h.redo = function() {
                if (this.historyIndex < this.history.length - 1) var a = this.history[++this.historyIndex];
                if (a) {
                    this.ignoreHistory = !0;
                    var d = a.worksheet;
                    d.openWorksheet();
                    b.call(this, a);
                    if (a.cascade && a.cascade.length)
                        for (var e = 0; e < a.cascade.length; e++) b.call(this, a.cascade[e]);
                    a.selection && d.updateSelectionFromCoords(a.selection);
                    this.ignoreHistory = !1
                }
            };
            h.reset = function() {
                this.historyIndex = -1;
                this.history = []
            };
            return h
        }(),
        wa = function() {
            var h = function(c) {
                var b = null;
                if (Array.isArray(c)) b = c;
                else if (typeof c == _$_669e[109]) {
                    b = [];
                    var a = Object.keys(c);
                    for (var d = 0; d < a.length; d++) b.push({
                        name: a[d],
                        plugin: c[a[d]],
                        options: {}
                    })
                }
                if (b && b.length)
                    for (d =
                        0; d < b.length; d++) {
                        if (b[d].name && b[d].plugin) {
                            var e = b[d].name;
                            var g = b[d].plugin;
                            var f = b[d].options
                        }
                        typeof g === _$_669e[110] && (this.plugins[e] = g.call(r, this, f, this.config), typeof g.license == _$_669e[110] && r.license && g.license(r.license))
                    }
            };
            h.execute = function(c, b) {
                if (this.plugins) {
                    var a = Object.keys(this.plugins);
                    if (a.length)
                        for (var d = 0; d < a.length; d++)
                            if (typeof this.plugins[a[d]][c] == _$_669e[110]) {
                                var e = this.plugins[a[d]][c].apply(this.plugins[a[d]], b);
                                e && (c == _$_669e[111] && (b[4] = e), c == _$_669e[112] && (b[0] =
                                    e))
                            } return b
                }
            };
            h.onevent = function(c) {
                return function() {
                    var b, a, d = this.plugins;
                    if (d) {
                        var e = Object.keys(d);
                        if (e.length)
                            for (a = 0; a < e.length; a++)
                                if (typeof d[e[a]][c] == _$_669e[110]) {
                                    var g = d[e[a]][c].apply(d[e[a]], arguments);
                                    typeof g !== _$_669e[21] && (b = g)
                                }
                    }
                    return b
                }
            }(_$_669e[113]);
            return h
        }(),
        pa = function() {
            var h = function(c) {
                c.getEditor = h[_$_669e[56]];
                c.openEditor = h.open;
                c.closeEditor = h.close
            };
            h[_$_669e[56]] = function(c, b) {
                c = Ma[_$_669e[56]].call(this, c, b);
                return [c.type && typeof c.type == _$_669e[109] ? c.type : r.editors[c.type],
                    c
                ]
            };
            h.open = function(c, b, a) {
                if (!T.call(this.parent, this)) return !1;
                if (1 != c.classList.contains(_$_669e[26])) {
                    var d = parseInt(c.getAttribute(_$_669e[63])),
                        e = parseInt(c.getAttribute(_$_669e[59]));
                    if (!0 === this.options.locked) {
                        var g = H.getColumnNameFromCoords(d, e);
                        if (!this.options.cells || !this.options.cells[g] || !1 !== this.options.cells[g].locked) return !1
                    }
                    if (g = this.records[e][d].merged) d -= g[0], e -= g[1], c = this.records[e][d].element;
                    if (!U.attached.call(this, d, e)) {
                        if (this.results && 0 < this.results.length) return console.error(_$_669e[114]),
                            !1;
                        var f = g = null;
                        na.attached.call(this, d) || (g = d);
                        aa.attached.call(this, e) || (f = e);
                        J[_$_669e[62]].call(this, f, g)
                    }
                    this.resetBorders(_$_669e[115], !0);
                    h.position.call(this, d, e);
                    g = this.parent.input;
                    !a || 113 !== a.keyCode && a.type === _$_669e[116] || (g.selected = !0);
                    g.innerText = _$_669e[3];
                    G.call(this.parent, _$_669e[117], this, c, d, e);
                    b = 1 == b ? _$_669e[3] : R.call(this, d, e);
                    f = h[_$_669e[56]].call(this, d, e);
                    var k = !0;
                    typeof f[0].openEditor == _$_669e[110] && (!1 === f[0].openEditor(c, b, d, e, this, f[1], a) ? k = !1 : G.call(this.parent, _$_669e[118],
                        this, c, d, e, g, f[1]));
                    k && (1 == this.options.textOverflow && (this.records[e][d].element.style.overflow = _$_669e[119], 0 < d && this.records[e][d - 1].element && (this.records[e][d - 1].element.style.overflow = _$_669e[119])), this.edition = c, g.classList.add(_$_669e[120]), 1 == this.parent.config.editorFormulas && Z(b) && (g.classList.add(_$_669e[121]), ta.parse.call(this, g), c = document.createTextNode(_$_669e[3]), g.appendChild(c), jSuites.focus(g)))
                }
            };
            h.close = function(c, b) {
                null === c && (c = this.edition);
                this.parent.input.classList.contains(_$_669e[121]) &&
                    ta.close.call(this, this.parent.input);
                var a, d = parseInt(c.getAttribute(_$_669e[63])),
                    e = parseInt(c.getAttribute(_$_669e[59]));
                1 == this.options.textOverflow && (a = this.records[e][d + 1]) && a.element && a.element.innerText == _$_669e[3] && (this.records[e][d].element.style.overflow = _$_669e[3]);
                a = h[_$_669e[56]].call(this, d, e);
                if (1 == b) {
                    var g = null;
                    typeof a[0].closeEditor == _$_669e[110] && (a = a[0].closeEditor(c, !0, d, e, this, a[1]), void 0 !== a && (g = a));
                    null !== g && R.call(this, d, e) != g && this.setValue(c, g)
                } else a[0].closeEditor(c,
                    !1, d, e, this, a[1]);
                G.call(this.parent, _$_669e[122], this, c, d, e, g, b);
                a = this.parent.input;
                a.selected = !1;
                a.onblur && (a.onblur = null);
                a.children[0] && a.children[0].onblur && (a.children[0].onblur = null);
                a.update && (a.update = null);
                pa.build(this.parent);
                a.removeAttribute(_$_669e[123]);
                a.removeAttribute(_$_669e[124]);
                a.removeAttribute(_$_669e[125]);
                this.edition = a.mask = null;
                this.refreshBorders()
            };
            h.position = function(c, b) {
                var a = null;
                if (a = this.records[b][c].element) {
                    var d = this.parent.element.getBoundingClientRect();
                    a = a.getBoundingClientRect();
                    var e = this.parent.input;
                    e.x = c;
                    e.y = b;
                    0 == a.width && 0 == a.height ? e.style.opacity = _$_669e[126] : (e.setAttribute(_$_669e[127], _$_669e[3]), c = a.top - d.top + 1, b = a.left - d.left + 1, e.style.top = c + _$_669e[57], e.style.left = b + _$_669e[57], e.style.minWidth = a.width - 1 + _$_669e[57], e.style.minHeight = a.height - 1 + _$_669e[57], setTimeout(function() {
                        e.focus()
                    }, 0))
                }
            };
            h.build = function(c) {
                c.input || (c.input = document.createElement(_$_669e[32]));
                c.input.className = _$_669e[128];
                c.input.setAttribute(_$_669e[129],
                    !0);
                c.input.innerText = _$_669e[3];
                c.input.oninput = function(b) {
                    r.current && (1 == r.current.parent.config.editorFormulas && (Z(b.target.innerText) ? (b.target.classList.add(_$_669e[121]), b.target.selected = !0) : b.target.classList.contains(_$_669e[121]) && ta.close.call(r.current, this)), G.call(r.current.parent, _$_669e[130], r.current, b))
                };
                c.input.onclick = function() {
                    this.selected = !0
                };
                return c.input
            };
            return h
        }(),
        G = function(h) {
            var c;
            if (!this.ignoreEvents) {
                if (typeof this.config.onevent == _$_669e[110]) {
                    var b = this.config.onevent.apply(this,
                        arguments);
                    typeof b !== _$_669e[21] && (c = b)
                }
                typeof this.config[h] == _$_669e[110] && (b = this.config[h].apply(this, Array.prototype.slice.call(arguments, 1)), typeof b !== _$_669e[21] && (c = b));
                b = wa.onevent.apply(this, arguments);
                typeof b !== _$_669e[21] && (c = b)
            }
            return c
        },
        za = function() {
            var h = function(c) {
                c.sequence = 0;
                c.getNextSequence = h.next;
                c.getRowId = h.getId;
                c.setRowId = h.setId;
                c.getRowById = h.getRowById
            };
            h[_$_669e[56]] = function() {
                return this.sequence
            };
            h[_$_669e[11]] = function(c, b) {
                if (c > this.sequence || b) this.sequence =
                    c
            };
            h.next = function() {
                return 1 == this.options.sequence ? ++this.sequence : null
            };
            h.reset = function() {
                this.sequence = 0
            };
            h.getId = function(c) {
                if (this.rows[c] && this.rows[c].id) return this.rows[c].id;
                var b = this.getPrimaryKey();
                if (!1 !== b) return R.call(this, b, c)
            };
            h.setId = function(c, b) {
                if (void 0 !== b && 0 <= parseInt(c)) {
                    var a = {};
                    a[c] = b
                } else a = c;
                c = Object.keys(a);
                if (c.length) {
                    for (var d = 0; d < c.length; d++) b = a[c[d]], this.rows[c[d]].id = b, h[_$_669e[11]].call(this, b);
                    L.call(this, _$_669e[131], [a]);
                    G.call(this.parent, _$_669e[132],
                        this, a)
                }
            };
            h.getRowById = function(c, b) {
                for (var a = 0; a < this.rows.length; a++)
                    if (h.getId.call(this, a) == c) return !0 === b ? this.rows[a] : this.options.data[a];
                return !1
            };
            return h
        }(),
        Wa = function() {
            var h = function(b) {
                    for (var a = [], d = {}, e = 0; e < b.length; e++) {
                        var g = b[e].x,
                            f = b[e].y,
                            k = za.getId.call(this, f);
                        if (!k)
                            if (d[f]) k = d[f];
                            else if (k = za.next.call(this)) d[f] = k;
                        a[f] || (a[f] = {
                            id: k,
                            row: f,
                            data: {}
                        });
                        g = ra.call(this, g);
                        a[f].data[g] = b[e].value
                    }
                    za.setId.call(this, d);
                    return a.filter(function(l) {
                        return null != l
                    })
                },
                c = function(b) {
                    var a =
                        b[0];
                    if (0 <= [_$_669e[133], _$_669e[134], _$_669e[131], _$_669e[135]].indexOf(a)) return !1;
                    var d = {};
                    a == _$_669e[97] || a == _$_669e[136] ? (a = _$_669e[137], d[a] = h.call(this, b[1].data)) : Array.isArray(b[1]) && 1 == b[1].length ? d[a] = b[1][0] : d[a] = b[1];
                    return d
                };
            return function(b, a, d, e) {
                var g = this.parent,
                    f = this,
                    k = c.call(this, a);
                if (k) a = k;
                else if (!1 === k) return !1;
                if (k = G.call(g, _$_669e[138], g, f, a)) a = k;
                else if (!1 === k) return !1;
                Qa.call(g, !0);
                jSuites.ajax({
                    url: b,
                    method: _$_669e[139],
                    dataType: _$_669e[73],
                    data: {
                        data: JSON.stringify(a)
                    },
                    queue: !0,
                    beforeSend: function(l) {
                        d && l.setRequestHeader(_$_669e[140], _$_669e[141] + d);
                        G.call(g, _$_669e[142], f, l)
                    },
                    success: function(l) {
                        g.element.classList.contains(_$_669e[143]) && g.element.classList.remove(_$_669e[143]);
                        l && (l.message && (l.name = g.config.application, jSuites.notification.isVisible() ? console.log(l.name + _$_669e[144] + l.message) : jSuites.notification(l)), l.success ? typeof e === _$_669e[110] ? e(l) : l.data && Array.isArray(l.data) && U.updateAll.call(f, l.data) : l.error && (alert(_$_669e[145]), window.open(window.location.href,
                            _$_669e[146])));
                        G.call(g, _$_669e[147], g, f, a, l)
                    },
                    error: function() {
                        g.element.classList.add(_$_669e[148])
                    }
                })
            }
        }(),
        L = function() {
            return function() {
                wa.execute.call(this.parent, _$_669e[149], [this, arguments[0], arguments[1]]);
                if (!this.parent.ignorePersistence) {
                    var h = null;
                    if (this.parent.config.server) {
                        var c = this.parent.worksheets.indexOf(this);
                        h = this.parent.config.server;
                        h = -1 == h.indexOf(_$_669e[150]) ? h + (_$_669e[151] + c) : h + (_$_669e[152] + c)
                    } else this.options.persistence && (h = typeof this.options.persistence == _$_669e[25] ?
                        this.options.persistence : this.options.url);
                    h && Wa.call(this, h, arguments)
                }
            }
        }(),
        Ia = function(h) {
            if (this.selectedCell) {
                var c = this.selectedCell[0],
                    b = this.selectedCell[1];
                this.records && this.records[b] && this.records[b][c] && this.records[b][c].element || (this.selectedCell = this.getHighlighted(), c = this.selectedCell[0], b = this.selectedCell[1]);
                var a = this.getHighlighted(),
                    d = a[0],
                    e = a[1],
                    g = a[2];
                a = a[3];
                var f = h ? _$_669e[153] : _$_669e[154];
                this.records[b] && this.records[b][c] && this.records[b][c].element && (this.records[b][c].element.classList[f](_$_669e[155]),
                    this.cursor = h ? this.records[b][c].element : null);
                for (h = d; h <= g; h++)
                    if (this.headers[h]) this.headers[h].classList[f](_$_669e[156]);
                for (; e <= a; e++)
                    if (this.rows[e] && this.rows[e].element) this.rows[e].element.classList[f](_$_669e[156])
            }
        },
        ka = function() {
            var h = function(c) {
                c.resetSelection = h.reset;
                c.updateSelection = h.fromElements;
                c.updateSelectionFromCoords = h[_$_669e[11]];
                c.selectAll = h.all;
                c.isSelected = h.isSelected;
                c.getHighlighted = h.getHighlighted;
                c.getRange = h.getRange;
                c.getSelection = h[_$_669e[56]]
            };
            h[_$_669e[56]] =
                function() {
                    var c = this.selectedCell;
                    if (c[0] < c[2]) var b = c[0],
                        a = c[2];
                    else b = c[2], a = c[0];
                    if (c[1] < c[3]) {
                        var d = c[1];
                        c = c[3]
                    } else d = c[3], c = c[1];
                    return [b, d, a, c]
                };
            h.getHighlighted = function() {
                var c = this.borders;
                return c && c.main && null !== c.main.x1 ? (c = c.main, [c.x1, c.y1, c.x2, c.y2]) : null
            };
            h.isSelected = function(c, b, a) {
                if (!a && (a = this.getHighlighted(), !a || null === a[0])) return !1;
                var d = a[0],
                    e = a[1],
                    g = a[2];
                a = a[3];
                g || (g = d);
                a || (a = e);
                return null == c ? b >= e && b <= a : null == b ? c >= d && c <= g : c >= d && c <= g && b >= e && b <= a
            };
            h.reset = function() {
                if (this.selectedCell) c =
                    1, this.edition && pa.close.call(this, this.edition, !0), this.resetBorders(_$_669e[157], !0), this.resetBorders(_$_669e[115], !0), this.selectedCell = null, 1 == c && G.call(this.parent, _$_669e[158], this);
                else var c = 0;
                return c
            };
            h.refresh = function() {
                this.selectedCell && h[_$_669e[11]].call(this, this.selectedCell)
            };
            h.isValid = function(c, b, a, d) {
                if (!this) return !1;
                Array.isArray(c) && (d = c[3], a = c[2], b = c[1], c = c[0]);
                return !(c >= this.options.columns.length || b >= this.rows.length || a >= this.options.columns.length || d >= this.rows.length)
            };
            h[_$_669e[11]] = function(c, b, a, d, e, g, f) {
                g || (g = _$_669e[157]);
                Array.isArray(c) && (d = c[3], a = c[2], b = c[1], c = c[0]);
                null == a && (a = c);
                null == d && (d = b);
                if (!h.isValid.call(this, c, b, a, d)) return !1;
                g == _$_669e[157] && r.current != this && this.openWorksheet();
                if (null != c) {
                    if (parseInt(c) < parseInt(a)) var k = parseInt(c),
                        l = parseInt(a);
                    else k = parseInt(a), l = parseInt(c);
                    if (parseInt(b) < parseInt(d)) var n = parseInt(b),
                        p = parseInt(d);
                    else n = parseInt(d), p = parseInt(b);
                    var q;
                    var v = {};
                    if (this.options.mergeCells) {
                        for (var w = n; w <= p; w++)
                            for (var x =
                                    k; x <= l; x++)
                                if (q = this.records[w][x].merged) q = H.getColumnNameFromCoords(x - q[0], w - q[1]), v[q] = !0;
                        w = Object.keys(v);
                        if (w.length)
                            for (x = 0; x < w.length; x++) q = H.getCoordsFromColumnName(w[x]), v = this.options.mergeCells[w[x]], q[0] < k && (k = q[0]), q[0] + v[0] - 1 > l && (l = q[0] + v[0] - 1), q[1] < n && (n = q[1]), q[1] + v[1] - 1 > p && (p = q[1] + v[1] - 1)
                    }
                }
                g == _$_669e[157] && (this.selectedCell ? (Ia.call(this, !1), q = 1) : q = 0, this.selectedCell = [c, b, a, d], this.records[d][a].element && e && (x = e.target.getAttribute(_$_669e[63]), v = e.target.getAttribute(_$_669e[59]),
                    null !== x && null !== v && (J.adjust.call(this, a, d, 3), J.adjust.call(this, a, d, 2), J.adjust.call(this, a, d, 1), J.adjust.call(this, a, d, 0))), pa.position.call(this, c, b), 0 == q && G.call(this.parent, _$_669e[159], this));
                ca[_$_669e[11]].call(this, k, n, l, p, g, f);
                g == _$_669e[157] && (G.call(this.parent, _$_669e[160], this, k, n, l, p, e), ma.update.call(this.parent, this, this.selectedCell))
            };
            h.fromElements = function(c, b, a) {
                var d = c.getAttribute(_$_669e[63]);
                c = c.getAttribute(_$_669e[59]);
                if (b) {
                    var e = b.getAttribute(_$_669e[63]);
                    b = b.getAttribute(_$_669e[59])
                } else e =
                    d, b = c;
                h[_$_669e[11]].call(this, d, c, e, b, a)
            };
            h.all = function() {
                h[_$_669e[11]].call(this, 0, 0, this.options.columns.length - 1, this.records.length - 1)
            };
            h.getRange = function() {
                var c = this.selectedCell;
                if (!c) return _$_669e[3];
                var b = H.getColumnNameFromCoords(c[0], c[1]);
                c = H.getColumnNameFromCoords(c[2], c[3]);
                var a = this.options.worksheetName;
                a = -1 != a.indexOf(_$_669e[161]) ? _$_669e[162] + a + _$_669e[163] : a + _$_669e[19];
                return b == c ? a + b : a + H.getRangeFromTokens(r.helpers.getTokensFromRange(b + _$_669e[164] + c))
            };
            return h
        }(),
        ca =
        function() {
            var h = function(a) {
                a.borders = [];
                a.setBorder = h[_$_669e[11]];
                a.getBorder = h[_$_669e[56]];
                a.resetBorders = h.reset;
                a.refreshBorders = h.refresh
            };
            h[_$_669e[56]] = function(a) {
                return this.borders[a]
            };
            var c = function(a, d, e, g, f, k) {
                    if (!ka.isValid.call(this, a, d, e, g)) return !1;
                    var l = !1,
                        n = null,
                        p = null,
                        q = null,
                        v = null,
                        w = null,
                        x = this.options.freezeColumns || 0,
                        y = this.options.freezeRows || 0;
                    f || (f = 0);
                    this.borders[f] ? n = this.borders[f] : (n = {}, n.element = document.createElement(_$_669e[32]), n.element.classList.add(_$_669e[165]),
                        f == _$_669e[115] || f == _$_669e[166] || f == _$_669e[157] ? (n.element.classList.add(_$_669e[167] + f), n.color = _$_669e[3]) : (n.color = k ? k : jSuites.randomColor(!0), n.element.style.backgroundColor = k + _$_669e[168], n.element.style.borderColor = k), f == _$_669e[157] && (k = function(F, u) {
                                n[F] = document.createElement(_$_669e[32]);
                                n[F].classList.add(_$_669e[169]);
                                n[F].setAttribute(_$_669e[170], u);
                                n.element.appendChild(n[F]);
                                !1 === _$_669e[171] in document.documentElement && (n[F].style.display = _$_669e[52])
                            }, k(_$_669e[172], 0), n.moveLeft.style.left =
                            _$_669e[173], n.moveLeft.style.marginTop = _$_669e[174], n.moveLeft.style.marginLeft = _$_669e[174], k(_$_669e[175], 1), n.moveTop.style.left = _$_669e[173], n.moveTop.style.marginTop = _$_669e[174], n.moveTop.style.marginLeft = _$_669e[174], k(_$_669e[176], 2), n.moveRight.style.right = _$_669e[173], n.moveRight.style.marginTop = _$_669e[174], n.moveRight.style.marginRight = _$_669e[174], k(_$_669e[177], 3), n.moveDown.style.bottom = _$_669e[173], n.moveDown.style.marginLeft = _$_669e[174], n.moveDown.style.marginBottom = _$_669e[174]),
                        this.content.appendChild(n.element));
                    n.x1 = a;
                    n.y1 = d;
                    n.x2 = e;
                    n.y2 = g;
                    n.element.style.top = _$_669e[178];
                    n.element.style.left = _$_669e[178];
                    n.element.style.borderTopColor = n.color;
                    n.element.style.borderBottomColor = n.color;
                    n.element.style.borderLeftColor = n.color;
                    n.element.style.borderRightColor = n.color;
                    n.element.style.backgroundPositionX = _$_669e[3];
                    n.element.style.backgroundPositionY = _$_669e[3];
                    n.active = 0;
                    k = this.visibleRows[y];
                    var B = this.visibleRows[this.visibleRows.length - 1],
                        A = this.visibleCols[x];
                    a > this.visibleCols[this.visibleCols.length -
                        1] || d > B || e < A && a >= x || g < k && d >= y ? l = !0 : (!1 !== this.options.columns[a].visible && this.headers[a] && this.headers[a].offsetWidth || (a = Ua.call(this, a, d), n.element.style.borderLeftColor = _$_669e[179], n.element.style.backgroundPositionX = _$_669e[33]), !1 !== this.options.columns[e].visible && this.headers[e] && this.headers[e].offsetWidth || (e = Va.call(this, e, g), n.element.style.borderRightColor = _$_669e[179], n.element.style.backgroundPositionX = _$_669e[180]), p = this.headers[a] && this.headers[a].offsetLeft ? this.headers[a].offsetLeft :
                        0, this.headers[e] && this.headers[e].parentNode ? (q = this.headers[e].offsetLeft, q = q - p + parseInt(this.options.columns[e].width)) : (v = this.thead.lastChild.lastChild.getAttribute(_$_669e[63]), x && e < v ? (q = this.thead.lastChild.children[x - 1].offsetLeft, q = q - p + parseInt(this.options.columns[x - 1].width)) : (q = this.thead.lastChild.lastChild.offsetLeft, q = q - p + parseInt(this.options.columns[v].width)), n.element.style.borderRightColor = _$_669e[179], n.element.style.backgroundPositionX = _$_669e[180]), p += this.table.offsetLeft,
                        !1 !== this.rows[d].visible && this.rows[d].element && this.rows[d].element.offsetHeight || (d = Ta.call(this, a, d), n.element.style.borderTopColor = _$_669e[179], n.element.style.backgroundPositionY = _$_669e[181]), !1 !== this.rows[g].visible && this.rows[g].element && this.rows[g].element.offsetHeight || (g = Sa.call(this, e, g), n.element.style.borderBottomColor = _$_669e[179], n.element.style.backgroundPositionY = _$_669e[182]), v = this.rows[d].element && this.rows[d].element.offsetTop ? this.rows[d].element.firstChild.offsetTop : 0, this.rows[g].element &&
                        this.rows[g].element.offsetTop ? (w = this.rows[g].element.firstChild.offsetTop, w = w - v + this.rows[g].element.firstChild.offsetHeight) : (this.tbody.lastChild ? (w = this.tbody.lastChild.getAttribute(_$_669e[59]), y && g < w ? (w = this.tbody.children[y - 1].firstChild.offsetTop, w = w - v + this.tbody.children[y - 1].firstChild.offsetHeight) : (w = this.tbody.lastChild.firstChild.offsetTop + this.tbody.lastChild.firstChild.offsetHeight - (this.thead.offsetHeight - 3), w = w - v + this.tbody.lastChild.firstChild.offsetHeight)) : w = 0, n.element.style.borderBottomColor =
                            _$_669e[179], n.element.style.backgroundPositionY = _$_669e[182]), v += this.table.offsetTop, q++, w++, this.options.freezeColumns && (k = ea.getWidth.call(this), a = a < x, e = e < x, a && e || (a && !e ? (e = this.content.scrollLeft + k, e > p + q && (q = q + (e - (p + q)) - 1, n.element.style.borderRightColor = _$_669e[179], n.element.style.backgroundPositionX = _$_669e[180])) : (e = this.content.scrollLeft + k, e > p && (e -= p, e > q ? p = v = -2E3 : (p += e, q -= e, n.element.style.borderLeftColor = _$_669e[179], n.element.style.backgroundPositionX = _$_669e[33]))))), this.options.freezeRows &&
                        (k = ja.getHeight.call(this), a = d < y, e = g < y, a && e || (a && !e ? (d = this.content.scrollTop + k, d > v + w && (w = w + (d - (v + w)) - 1, n.element.style.borderBottomColor = _$_669e[179], n.element.style.backgroundPositionY = _$_669e[182])) : (d = this.content.scrollTop + k, d > v && (e = d - v, e > w ? p = v = -2E3 : (v += e, w -= e, n.element.style.borderTopColor = _$_669e[179], n.element.style.backgroundPositionY = _$_669e[181]))))), n.element.style.top = v + _$_669e[57], n.element.style.left = p + _$_669e[57], n.element.style.width = q + _$_669e[57], n.element.style.height = w + _$_669e[57],
                        n.t = v, n.l = p, n.w = q, n.h = w, f == _$_669e[157] && (n.moveLeft.style.top = n.h / 2 + _$_669e[57], n.moveRight.style.top = n.h / 2 + _$_669e[57], n.moveTop.style.left = n.w / 2 + _$_669e[57], n.moveDown.style.left = n.w / 2 + _$_669e[57]), n.active = 1);
                    f == _$_669e[157] && (l || n.element.style.borderBottomColor || n.element.style.borderRightColor ? (this.corner.style.top = _$_669e[178], this.corner.style.left = _$_669e[178]) : (this.corner.style.top = v + w - 4 + _$_669e[57], this.corner.style.left = p + q - 4 + _$_669e[57]));
                    this.borders[f] = n
                },
                b = function(a, d) {
                    a.element.style.top =
                        _$_669e[178];
                    a.element.style.left = _$_669e[178];
                    a.active = 0;
                    d && (a.x1 = null, a.y1 = null, a.x2 = null, a.y2 = null)
                };
            h[_$_669e[11]] = function(a, d, e, g, f, k) {
                c.call(this, a, d, e, g, f, k);
                f == _$_669e[157] && (Ia.call(this, !0), L.call(this, _$_669e[133], {
                    x1: a,
                    y1: d,
                    x2: e,
                    y2: g
                }))
            };
            h.reset = function(a, d) {
                a && a != _$_669e[157] || (this.corner.style.top = _$_669e[178], this.corner.style.left = _$_669e[178], Ia.call(this, !1));
                if (a) this.borders[a] && b(this.borders[a], d);
                else
                    for (var e = Object.keys(this.borders), g = 0; g < e.length; g++) b(this.borders[e[g]],
                        d);
                a && a != _$_669e[157] || L.call(this, _$_669e[134])
            };
            h.refresh = function(a) {
                var d = this.borders;
                if (a) d[a] && null != d[a].x1 && c.call(this, d[a].x1, d[a].y1, d[a].x2, d[a].y2, a);
                else {
                    a = Object.keys(d);
                    for (var e = 0; e < a.length; e++) d[a[e]] && null != d[a[e]].x1 && c.call(this, d[a[e]].x1, d[a[e]].y1, d[a[e]].x2, d[a[e]].y2, a[e])
                }
                this.edition && (d = this.parent.input, pa.position.call(this, d.x, d.y))
            };
            h.destroy = function(a) {
                var d;
                this && (d = this.borders[a]) && (d.element.remove(), delete this.borders[a])
            };
            return h
        }(),
        Ja = function() {
            var h =
                function(c) {
                    c.setRowGroup = h[_$_669e[11]];
                    c.openRowGroup = h.open;
                    c.closeRowGroup = h.close;
                    c.resetRowGroup = h.reset
                };
            h[_$_669e[11]] = function(c, b) {
                if (Array.isArray(b)) var a = b;
                else {
                    a = [];
                    for (var d = 1; d <= b; d++) a.push(c + d)
                }
                this.options.rows[c] || (this.options.rows[c] = {});
                this.options.rows[c].group = a;
                d = this.rows[c];
                d.group = a;
                d.state = !1;
                aa.visibility.call(this, a, !1);
                h.icon.call(this, c);
                Q.call(this.parent, {
                    worksheet: this,
                    action: _$_669e[103],
                    row: c,
                    elements: a
                });
                L.call(this, _$_669e[103], {
                    row: c,
                    elements: b
                });
                G.call(this.parent,
                    _$_669e[190], this, c, b)
            };
            h.icon = function(c) {
                var b = null,
                    a = this,
                    d = this.rows[c];
                d.element && (b = d.element.firstChild.children[0], b || (b = document.createElement(_$_669e[191]), b.classList.add(_$_669e[192]), b.onclick = function() {
                    d.state ? h.close.call(a, c) : h.open.call(a, c)
                }, d.element.firstChild.appendChild(b)), b.innerText = d.state ? _$_669e[154] : _$_669e[153])
            };
            h.setState = function(c, b) {
                var a = this.rows[c];
                a && a.group && (a.state = b, aa.visibility.call(this, a.group, b), h.icon.call(this, c), Q.call(this.parent, {
                    worksheet: this,
                    action: b ? _$_669e[105] : _$_669e[106],
                    row: c
                }), L.call(this, b ? _$_669e[105] : _$_669e[106]), G.call(this.parent, b ? _$_669e[193] : _$_669e[194], this, c))
            };
            h.open = function(c) {
                h.setState.call(this, c, !0)
            };
            h.close = function(c) {
                h.setState.call(this, c, !1)
            };
            h.build = function() {
                var c = this.options.rows,
                    b, a, d, e = Object.keys(c);
                if (e.length)
                    for (d = 0; d < e.length; d++)
                        if (!c[e[d]].state && (b = c[e[d]].group) && b.length)
                            for (a = 0; a < b.length; a++) c[b[a]] || (c[b[a]] = {}), c[b[a]].visible = !1
            };
            h.reset = function(c) {
                var b = this.rows[c],
                    a;
                b.group && (b.state ||
                    h.open.call(this, c), b.element && (a = b.element.firstChild.children[0]) && a.remove(), Q.call(this.parent, {
                        worksheet: this,
                        action: _$_669e[104],
                        row: c,
                        elements: b.group
                    }), L.call(this, _$_669e[104], {
                        row: c
                    }), delete b.group, delete b.state, G.call(this.parent, _$_669e[195], this, c))
            };
            return h
        }(),
        Da = function() {
            var h = {
                start: function(c) {
                    if (!T.call(this.parent, this)) return !1;
                    if (null !== c.target.getAttribute(_$_669e[59])) {
                        var b = c.target.getAttribute(_$_669e[59]);
                        h.event = {
                            y: parseInt(b),
                            h: c.target.offsetHeight,
                            p: c.pageY
                        }
                    } else null !==
                        c.target.getAttribute(_$_669e[63]) && (b = c.target.getAttribute(_$_669e[63]), h.event = {
                            x: parseInt(b),
                            w: c.target.offsetWidth,
                            p: c.pageX
                        })
                },
                end: function(c) {
                    var b = [];
                    if (null != h.event.y) {
                        var a = parseInt(this.rows[h.event.y].element.offsetHeight);
                        var d = this.getSelectedRows(!0);
                        c = d.indexOf(parseInt(h.event.y));
                        if (d.length && -1 < c)
                            for (var e = 0; e < d.length; e++) b.push(aa.height.call(this, d[e]));
                        else d = []; - 1 == c ? (d.push(h.event.y), b.push(h.event.h)) : b[c] = h.event.h;
                        this.setHeight(d, a, b)
                    } else if (null != h.event.x && (a = parseInt(this.colgroup[h.event.x].getAttribute(_$_669e[196])),
                            h.event.w !== a)) {
                        d = this.getSelectedColumns();
                        c = d.indexOf(parseInt(h.event.x));
                        if (d.length && -1 < c)
                            for (e = 0; e < d.length; e++) b.push(parseInt(this.options.columns[d[e]].width));
                        else d = []; - 1 == c ? (d.push(h.event.x), b.push(h.event.w)) : b[c] = h.event.w;
                        this.setWidth(d, a, b)
                    }
                    h.event = null
                },
                update: function(c) {
                    null != h.event.y ? (c = c.pageY - h.event.p, 0 < h.event.h + c && (this.rows[h.event.y].element.style.height = h.event.h + c + _$_669e[57], this.refreshBorders())) : null != h.event.x && (c = c.pageX - h.event.p, 0 < h.event.w + c && (this.colgroup[h.event.x].setAttribute(_$_669e[196],
                        h.event.w + c), this.options.columns[h.event.x].width = h.event.w + c), 0 < this.options.freezeColumns && ea.refresh.call(this), this.refreshBorders())
                },
                cancel: function() {
                    h.end.call(this)
                }
            };
            return h
        }(),
        ya = function() {
            var h = {
                start: function(c) {
                    if (T.call(this.parent, this)) {
                        this.resetSelection();
                        var b = this.parent.helper;
                        if (null !== c.target.getAttribute(_$_669e[59])) {
                            var a = parseInt(c.target.getAttribute(_$_669e[59]));
                            this.merged.rows[a] ? console.error(_$_669e[197]) : this.results ? console.error(_$_669e[198]) : (h.event = {
                                y: a,
                                d: a
                            }, this.rows[a].element.classList.add(_$_669e[199]), c = this.parent.element.getBoundingClientRect(), a = this.rows[a].element.getBoundingClientRect(), b.style.display = _$_669e[24], b.style.top = a.top - c.top + _$_669e[57], b.style.left = _$_669e[200], b.style.width = this.content.offsetWidth + _$_669e[57], b.style.height = _$_669e[200], b.classList.add(_$_669e[201]))
                        } else if (null !== c.target.getAttribute(_$_669e[63]))
                            if (a = parseInt(c.target.getAttribute(_$_669e[63])), this.merged.cols[a]) console.error(_$_669e[202]);
                            else {
                                h.event = {
                                    x: a,
                                    d: a
                                };
                                this.headers[a].classList.add(_$_669e[199]);
                                for (c = 0; c < this.tbody.children.length; c++) {
                                    var d = this.tbody.children[c].getAttribute(_$_669e[59]);
                                    this.records[d][a].element.classList.add(_$_669e[199])
                                }
                                c = this.parent.element.getBoundingClientRect();
                                a = this.headers[a].getBoundingClientRect();
                                d = this.content.getBoundingClientRect();
                                b.style.display = _$_669e[24];
                                b.style.top = d.top - c.top + _$_669e[57];
                                b.style.left = a.left - d.left + _$_669e[57];
                                b.style.width = _$_669e[200];
                                b.style.height = d.height + _$_669e[57];
                                b.classList.add(_$_669e[203])
                            }
                    } else return !1
                },
                end: function() {
                    if (null != h.event.y) this.rows[h.event.y].element.classList.remove(_$_669e[199]), null != h.event.d && h.event.y != h.event.d && (this.moveRow(h.event.y, h.event.d), this.updateSelectionFromCoords(0, h.event.d, this.options.columns.length - 1, h.event.d)), this.parent.helper.classList.remove(_$_669e[201]);
                    else if (null != h.event.x) {
                        this.headers[h.event.x].classList.remove(_$_669e[199]);
                        for (var c = 0; c < this.tbody.children.length; c++) {
                            var b = this.tbody.children[c].getAttribute(_$_669e[59]);
                            this.records[b][h.event.x].element.classList.remove(_$_669e[199])
                        }
                        null !=
                            h.event.d && h.event.x != h.event.d && (this.moveColumn(h.event.x, h.event.d), this.refreshBorders());
                        this.parent.helper.classList.remove(_$_669e[203])
                    }
                    this.parent.helper.style.display = _$_669e[3];
                    h.event = null
                },
                cancel: function() {
                    h.event.d = null;
                    h.end.call(this)
                },
                update: function(c) {
                    if (null != h.event.y) {
                        var b = c.target.getAttribute(_$_669e[59]);
                        if (null != b) {
                            var a = b = parseInt(b),
                                d = c.target.clientHeight / 2 > c.offsetY;
                            d ? h.event.y < b && (a = b - 1) : h.event.y > b && (a = b + 1);
                            c = !1;
                            this.merged.rows[a] && (h.event.y > a ? this.merged.rows[a -
                                1] && (c = !0) : this.merged.rows[a + 1] && (c = !0));
                            c || (h.event.d = a, a = this.parent.element.getBoundingClientRect(), b = this.rows[b].element.getBoundingClientRect(), this.parent.helper.style.top = (d ? b.top : b.bottom) - a.top + _$_669e[57])
                        }
                    } else d = c.target.getAttribute(_$_669e[63]), null != d && (a = d = parseInt(d), (b = c.target.clientWidth / 2 > c.offsetX) ? h.event.x < d && (a = d - 1) : h.event.x > d && (a = d + 1), c = !1, this.merged.cols[a] && (h.event.x > a ? this.merged.cols[a - 1] && (c = !0) : this.merged.cols[a + 1] && (c = !0)), c || (h.event.d = a, a = this.headers[d].getBoundingClientRect(),
                        d = this.content.getBoundingClientRect(), this.parent.helper.style.left = (b ? a.left : a.right) - d.left + _$_669e[57]))
                }
            };
            return h
        }(),
        sa = function() {
            var h = function(c) {
                c.getMerge = h[_$_669e[56]];
                c.setMerge = h[_$_669e[11]];
                c.updateMerge = h.update;
                c.removeMerge = h.remove;
                c.destroyMerged = c.destroyMerge = h.destroy
            };
            h.isMerged = function(c, b) {
                var a = H.getColumnNameFromCoords(c, b);
                if (a = this.merged.cells[a])
                    if (i = this.options.mergeCells[a]) m && void 0 !== m[a] || (i[1] > t && (t = i[1]), m && (m[a] = this.records[b][c].element && !this.records[b][c].merged ?
                        !0 : !1))
            };
            h.getAffected = function(c) {
                if (void 0 !== this[c]) {
                    for (var b = c, a = this[c], d = []; this[b] == a;) d.push(b), b--;
                    for (b = c + 1; this[b] == a;) d.push(b), b++;
                    d.sort(function(e, g) {
                        return e - g
                    });
                    return d
                }
                return !1
            };
            h[_$_669e[56]] = function(c) {
                return c ? this.options.mergeCells[c] : this.options.mergeCells
            };
            h[_$_669e[11]] = function(c, b, a) {
                if (!T.call(this.parent, this)) return !1;
                this.options.mergeCells || (this.options.mergeCells = {});
                var d, e = null,
                    g = {},
                    f = {},
                    k = {},
                    l = !1;
                typeof c == _$_669e[25] ? k[c] = [b, a] : typeof c == _$_669e[109] ? k = c :
                    ((e = this.getHighlighted()) ? e = [H.getColumnNameFromCoords(e[0], e[1]), e[2] - e[0] + 1, e[3] - e[1] + 1] : (jSuites.notification({
                        message: D(_$_669e[204])
                    }), e = !1), e && (k[e[0]] = [e[1], e[2]]));
                b = Object.keys(k);
                if (b.length) {
                    for (a = 0; a < b.length; a++) {
                        var n = b[a],
                            p = k[b[a]][0],
                            q = k[b[a]][1];
                        p = parseInt(p);
                        q = parseInt(q);
                        if ((!p || 2 > p) && (!q || 2 > q)) console.log(D(_$_669e[205]) + _$_669e[164] + n);
                        else if (this.options.mergeCells[n]) f[n] = [p, q], g[n] = this.options.mergeCells[n], h.update.call(this, n, p, q);
                        else {
                            e = H.getCoordsFromColumnName(n);
                            for (var v =
                                    e[1]; v < e[1] + q; v++)
                                for (var w = e[0]; w < e[0] + p; w++) c = H.getColumnNameFromCoords(w, v), (d = this.merged.cells[c]) && this.options.mergeCells[d] && (g[d] = this.options.mergeCells[d], h.applyDestroy.call(this, d));
                            f[n] = this.options.mergeCells[n] = [p, q];
                            h.applyCreate.call(this, n);
                            U.attached.call(this, e[0], e[1]) || (l = e)
                        }
                    }
                    Object.keys(f).length && (l && J[_$_669e[62]].call(this, e[1], e[0]), Q.call(this.parent, {
                            worksheet: this,
                            action: _$_669e[86],
                            newValue: f,
                            oldValue: g
                        }), L.call(this, _$_669e[86], [f]), h.build.call(this), ka.refresh.call(this),
                        ca.refresh.call(this), G.call(this.parent, _$_669e[206], this, f, g))
                }
            };
            h.update = function(c, b, a) {
                var d = this.options.mergeCells;
                if (d[c]) {
                    var e = d[c][1];
                    if (b !== d[c][0] || e !== a) h.applyDestroy.call(this, c), b && a && (d[c] = [b, a], h.applyCreate.call(this, c));
                    h.build.call(this)
                }
            };
            h.remove = function(c) {
                if (!T.call(this.parent, this)) return !1;
                var b = {},
                    a = {},
                    d = {};
                typeof c == _$_669e[25] ? d[c] = !0 : typeof c == _$_669e[109] && (d = c);
                c = Object.keys(d);
                for (d = 0; d < c.length; d++) this.options.mergeCells[c[d]] && (b[c[d]] = this.options.mergeCells[c[d]],
                    a[c[d]] = !0, h.applyDestroy.call(this, c[d], !0));
                h.build.call(this);
                Q.call(this.parent, {
                    worksheet: this,
                    action: _$_669e[87],
                    newValue: a,
                    oldValue: b
                });
                L.call(this, _$_669e[87], {
                    data: a
                })
            };
            h.destroy = function() {
                h.remove.call(this, this.options.mergeCells, !0)
            };
            h.applyCreate = function(c, b) {
                var a = this.options.mergeCells;
                if (a[c]) {
                    var d = a[c][0] || 1,
                        e = a[c][1] || 1;
                    c = H.getCoordsFromColumnName(c);
                    if (!this.records[c[1]][c[0]].merged || 1 == b)
                        if (a = U[_$_669e[56]].call(this, c[0], c[1])) {
                            a.classList.remove(_$_669e[65]);
                            a.style.display =
                                _$_669e[3];
                            1 < d ? a.setAttribute(_$_669e[207], d) : a.removeAttribute(_$_669e[207]);
                            1 < e ? a.setAttribute(_$_669e[208], e) : a.removeAttribute(_$_669e[208]);
                            1 == this.options.textOverflow && (a.style.overflow = _$_669e[119]);
                            for (var g = 0; g < e; g++)
                                for (var f = 0; f < d; f++) {
                                    b = c[0] + f;
                                    var k = c[1] + g;
                                    if (a = this.records[k][b])
                                        if (a.merged = [f, g, a.v, d, e], 0 != f || 0 != g) a.element || U[_$_669e[56]].call(this, b, k), a.element.classList.add(_$_669e[65]), a.merged[3] = a.element.innerText, a.element.innerText = _$_669e[3], R.call(this, b, k, _$_669e[3])
                                }
                        }
                }
            };
            h.applyDestroy = function(c, b) {
                var a = this.options.mergeCells;
                if (a[c]) {
                    var d = H.getCoordsFromColumnName(c),
                        e = d[0];
                    d = d[1];
                    var g = a[c][0],
                        f = a[c][1];
                    delete this.options.mergeCells[c];
                    for (c = 0; c < f; c++)
                        for (var k = 0; k < g; k++) this.records[d + c] && (a = this.records[d + c][e + k]) && a.element && (0 == k && 0 == c ? (a.element.removeAttribute(_$_669e[207]), a.element.removeAttribute(_$_669e[208])) : (a.element.classList.remove(_$_669e[65]), 1 == b && (R.call(this, e + k, d + c, a.merged[2]), a.element.innerHTML = a.merged[3]))), a.merged && delete a.merged
                }
            };
            h.updateConfig = function(c, b, a, d) {
                if (1 == b) var e = [a];
                else {
                    e = [];
                    for (var g = a; g < a + d; g++) e.push(g)
                }
                a = this.options.mergeCells;
                var f = Object.keys(a),
                    k = {};
                for (g = 0; g < f.length; g++) {
                    var l = H.getCoordsFromColumnName(f[g]);
                    var n = l[c];
                    for (l = n + a[f[g]][c] - b; n < l; n++) - 1 !== e.indexOf(n) && (k[f[g]] || (k[f[g]] = [a[f[g]][0], a[f[g]][1]], 1 == b && (k[f[g]][c] += d)), 0 == b && k[f[g]][c]--)
                }
                f = Object.keys(k);
                for (g = 0; g < f.length; g++) h.update.call(this, f[g], k[f[g]][0], k[f[g]][1]);
                h.build.call(this)
            };
            h.build = function() {
                var c = this.options.mergeCells;
                this.merged = {
                    cols: [],
                    rows: [],
                    cells: []
                };
                if (c)
                    for (var b = Object.keys(c), a = 0; a < b.length; a++) {
                        c[b[a]][0] = parseInt(c[b[a]][0]);
                        c[b[a]][1] = parseInt(c[b[a]][1]);
                        var d = H.getCoordsFromColumnName(b[a]),
                            e = parseInt(c[b[a]][0]),
                            g = parseInt(c[b[a]][1]);
                        if (1 < e)
                            for (var f = 0; f < e; f++) this.merged.cols[d[0] + f] = !0;
                        if (1 < g)
                            for (var k = 0; k < g; k++) this.merged.rows[d[1] + k] = !0;
                        for (k = 0; k < g; k++)
                            for (f = 0; f < e; f++) {
                                var l = H.getColumnNameFromCoords(d[0] + f, d[1] + k);
                                this.merged.cells[l] = b[a]
                            }
                    }
            };
            h.batch = function(c) {
                var b = Object.keys(c);
                if (b.length)
                    for (var a =
                            0; a < b.length; a++) !1 === c[b[a]] && h.applyCreate.call(this, b[a])
            };
            return h
        }(),
        Fa = function() {
            var h = function(c) {
                c.getHeader = h[_$_669e[56]];
                c.setHeader = h[_$_669e[11]];
                c.getHeaders = h.all
            };
            h[_$_669e[56]] = function(c) {
                var b;
                if (b = this.options.columns) return b[c].title || H.getColumnName(c)
            };
            h.all = function(c) {
                for (var b = [], a = 0; a < this.options.columns.length; a++) b.push(h[_$_669e[56]].call(this, a));
                return c ? b : b.join(this.options.csvDelimiter)
            };
            h[_$_669e[11]] = function(c, b) {
                if (!T.call(this.parent, this)) return !1;
                var a;
                if (a = this.options.columns[c]) {
                    var d = h[_$_669e[56]].call(this, c);
                    if (void 0 === b && (b = prompt(D(_$_669e[209]), d), !b)) return !1;
                    a.title = b;
                    if (a = this.headers[c]) b ? (a.innerHTML = b, a.setAttribute(_$_669e[210], b)) : (a.innerHTML = H.getColumnName(c), a.removeAttribute(_$_669e[210]));
                    Q.call(this.parent, {
                        worksheet: this,
                        action: _$_669e[92],
                        column: c,
                        oldValue: d,
                        newValue: b
                    });
                    L.call(this, _$_669e[92], {
                        column: c,
                        title: b
                    });
                    G.call(this.parent, _$_669e[211], this, c, b, d)
                }
            };
            h.create = function(c) {
                if (this.headers[c]) return this.headers[c];
                var b = document.createElement(_$_669e[185]);
                b.classList.add(_$_669e[212]);
                b.setAttribute(_$_669e[63], c);
                var a = document.createElement(_$_669e[213]);
                this.headers[c] = b;
                this.colgroup[c] = a;
                this.options.nestedHeaders && (b.style.top = this.thead.lastChild.offsetTop + _$_669e[57]);
                h.applyProperties.call(this, c);
                G.call(this.parent, _$_669e[214], this, c, b, this.options.columns[c]);
                return b
            };
            h.applyProperties = function(c) {
                var b = this.options.columns[c],
                    a = b.width || this.options.defaultColWidth,
                    d = this.headers[c];
                b.align &&
                    (d.style.textAlign = b.align);
                b.title ? (d.innerText = b.title, d.setAttribute(_$_669e[210], b.title)) : (d.innerText = H.getColumnName(c), d.removeAttribute(_$_669e[210]));
                b.tooltip && (d.title = b.tooltip);
                var e = this.options.filters;
                !0 === b.filter && (e = !0);
                !1 === b.filter && (e = !1);
                !0 === e && d.classList.add(_$_669e[215]);
                if (!1 === b.visible || b.type == _$_669e[119]) a = _$_669e[126], b.visible = !1;
                this.colgroup[c].setAttribute(_$_669e[196], a)
            };
            return h
        }(),
        ia = function() {
            var h = function(b, a, d) {
                    var e = this.footers.content,
                        g;
                    e && e[a] &&
                        e[a][b] && (g = e[a][b].element) && (void 0 === d && (e = this.options.footers, d = e[a] && e[a][b] ? e[a][b] : _$_669e[3]), Z(d) && 1 == this.parent.config.parseFormulas && (d = Y.execute.call(this, d, b, null, !1)), this.options.applyMaskOnFooters && (e = pa[_$_669e[56]].call(this, b), d = r.editors.general.parseValue(b, a, d, this, e[1], g)), d instanceof Element || d instanceof HTMLDocument ? (g.innerHTML = _$_669e[3], g.appendChild(d)) : 1 == this.parent.config.stripHTML ? g.innerText = d : g.innerHTML = d, G.call(this.parent, _$_669e[216], this, b, a, d, g))
                },
                c = function(b) {
                    b.getFooter =
                        c[_$_669e[56]];
                    b.setFooter = c[_$_669e[11]];
                    b.resetFooter = c.reset;
                    b.refreshFooter = c.refresh;
                    b.getFooterValue = c.value;
                    b.setFooterValue = c.value
                };
            c[_$_669e[56]] = function() {
                return this.options.footers
            };
            c[_$_669e[11]] = function(b) {
                if (b) {
                    var a = this.options.footers;
                    this.options.footers = b
                }
                this.options.footers != a && (ia.build.call(this), c.render.call(this), L.call(this, _$_669e[217], {
                    data: this.options.footers
                }), G.call(this.parent, _$_669e[218], this, this.options.footers, a))
            };
            c.refresh = function() {
                var b = this.options.footers;
                if (b && b.length)
                    for (var a = 0; a < b.length; a++)
                        for (var d = 0; d < b[a].length; d++) h.call(this, d, a, b[a][d])
            };
            c.value = function(b, a, d) {
                if (void 0 === d) return this.options.footers[a][b];
                h.call(this, b, a, d);
                L.call(this, _$_669e[219], {
                    x: b,
                    y: a,
                    value: d
                });
                G.call(this.parent, _$_669e[220], this, b, a, d)
            };
            c.reset = function() {
                var b = this.options.footers;
                this.options.footers = null;
                this.footers.content = null;
                this.tfoot.innerHTML = _$_669e[3];
                L.call(this, _$_669e[221], {});
                G.call(this.parent, _$_669e[218], this, null, b)
            };
            c.create = function(b,
                a) {
                var d = this.footers.content[a][b];
                if (d && d.element) return d.element;
                d = this.footers.content[a][b] = {};
                var e = this.options.columns[b].align;
                d.element = document.createElement(_$_669e[185]);
                d.element.setAttribute(_$_669e[63], b);
                d.element.setAttribute(_$_669e[59], a);
                e && (d.element.style.textAlign = e);
                h.call(this, b, a);
                return d.element
            };
            c.build = function() {
                this.footers = {
                    element: this.tfoot,
                    content: []
                };
                if (this.options.footers) {
                    this.tfoot.innerHTML = _$_669e[3];
                    for (var b = 0; b < this.options.footers.length; b++) {
                        if (!this.footers.content[b]) {
                            this.footers.content[b] = [];
                            var a = document.createElement(_$_669e[183]),
                                d = document.createElement(_$_669e[185]);
                            d.innerHTML = _$_669e[186];
                            a.appendChild(d);
                            this.footers.element.appendChild(a)
                        }
                        for (a = 0; a < this.options.columns.length; a++) this.options.footers[b][a] || (this.options.footers[b][a] = _$_669e[3]), this.footers.content[b][a] || (this.footers.content[b][a] = {})
                    }
                }
            };
            c.render = function() {
                for (var b = this.options.footers, a = 0; a < b.length; a++)
                    for (var d = 0; d < b[a].length; d++) {
                        var e = ia.create.call(this, d, a);
                        this.tfoot.children[a].appendChild(e)
                    }
            };
            c.adjust = function(b, a, d) {
                var e = [],
                    g = this.options.footers,
                    f = this.footers.content;
                if (g)
                    if (1 == d) {
                        if (this.headers[b] && this.headers[b].parentNode) var k = this.headers[b] == this.thead.lastChild.lastChild ? !0 : !1,
                            l = !0;
                        for (d = 0; d < g.length; d++) {
                            for (var n = g[d].splice(b), p = b; p < b + a; p++) g[d][p] = _$_669e[3];
                            g[d] = g[d].concat(n);
                            n = f[d].splice(b);
                            for (p = b; p < b + a; p++) f[d][p] = {
                                element: c.create.call(this, p, d)
                            }, l && (1 == k ? this.tfoot.children[d].appendChild(f[d][p].element) : this.tfoot.children[d].insertBefore(f[d][p].element, this.tfoot.children[d].children[b].nextSibling));
                            f[d] = f[d].concat(n)
                        }
                    } else
                        for (d = 0; d < g.length; d++) {
                            for (p = b; p < b + a; p++)(k = this.footers.content[d][p].element) && k.parentNode && k.parentNode.removeChild(k);
                            f[d].splice(b, a);
                            e[d] = g[d].splice(b, a)
                        }
                return e
            };
            return c
        }(),
        Aa = function() {
            var h = function(c) {
                c.setNestedCell = h.cell;
                c.getNestedCell = h.cell;
                c.setNestedHeaders = h[_$_669e[11]];
                c.getNestedHeaders = h[_$_669e[56]];
                c.resetNestedHeaders = h.reset
            };
            h[_$_669e[56]] = function() {
                return this.options.nestedHeaders
            };
            h[_$_669e[11]] = function(c) {
                h.reset.call(this);
                this.options.nestedHeaders =
                    c;
                h.build.call(this);
                J.call(this, !0, !0);
                ca.refresh.call(this);
                L.call(this, _$_669e[222], {
                    data: c
                });
                G.call(this.parent, _$_669e[223], this, c)
            };
            h.reset = function() {
                var c = [],
                    b = this.thead.children;
                this.options.nestedHeaders = null;
                for (var a = 0; a < b.length; a++) b[a].classList.contains(_$_669e[224]) && c.push(b[a]);
                if (c.length)
                    for (; b = c.shift();) b.remove();
                this.nested = null;
                J.call(this, !0, !0);
                ca.refresh.call(this);
                L.call(this, _$_669e[225]);
                G.call(this.parent, _$_669e[223], this, [])
            };
            h.cell = function(c, b, a) {
                if (a) {
                    if (!T.call(this.parent,
                            this)) return !1;
                    var d = this.options.nestedHeaders[b][c],
                        e = this.nested.content[b][c].element;
                    void 0 !== a.title && (e.innerText = a.title);
                    void 0 !== a.tooltip && (e.title = a.tooltip);
                    void 0 !== a.colspan && (d.colspan = a.colspan, e.setAttribute(_$_669e[207], a.colspan), e.style.display = 0 == d.colspan ? _$_669e[52] : _$_669e[3]);
                    L.call(this, _$_669e[226], {
                        x: c,
                        y: b,
                        properties: a
                    });
                    G.call(this.parent, _$_669e[227], this, c, b, a)
                } else return this.nested.content[b][c].element
            };
            h.range = function(c) {
                c = this.options.nestedHeaders[c];
                for (var b = [], a = 0, d = 0; d < c.length; d++) {
                    var e = parseInt(c[d].colspan);
                    b[d] = [a, a + e - 1];
                    a += e
                }
                return b
            };
            h.getColumns = function(c) {
                c = this.options.nestedHeaders[c];
                for (var b = [], a = 0; a < c.length; a++)
                    for (var d = c[a].colspan, e = 0; e < d; e++) b.push(a);
                return b
            };
            h.getAffected = function(c) {
                var b = [],
                    a = this.options.nestedHeaders;
                if (a) {
                    var d, e, g;
                    for (e = 0; e < a.length; e++)
                        if ((d = h.getColumns.call(this, e)) && void 0 !== d[c]) {
                            var f = h.range.call(this, e);
                            for (g = f[d[c]][0]; g <= f[d[c]][1]; g++) b.push(g)
                        }
                }
                return b
            };
            h.create = function(c, b) {
                if (this.nested.content[b][c]) return this.nested.content[b][c].element;
                var a = this.options.nestedHeaders[b][c] || {},
                    d = this.nested.content[b][c] = {};
                a.colspan || (a.colspan = 1);
                a.align || (a.align = _$_669e[228]);
                a.title || (a.title = _$_669e[3]);
                var e = document.createElement(_$_669e[185]);
                e.setAttribute(_$_669e[229], c);
                e.setAttribute(_$_669e[230], b);
                e.setAttribute(_$_669e[207], a.colspan);
                e.setAttribute(_$_669e[231], a.align);
                e.innerHTML = a.title;
                a.tooltip && (e.title = a.tooltip);
                return d.element = e
            };
            h.build = function() {
                this.nested || (this.nested = {
                    content: []
                });
                var c, b, a = 0;
                if (c = this.options.nestedHeaders) {
                    for (b =
                        0; b < c.length; b++) {
                        this.nested.content[b] = [];
                        var d = document.createElement(_$_669e[185]);
                        d.innerHTML = _$_669e[186];
                        var e = document.createElement(_$_669e[183]);
                        e.classList.add(_$_669e[224]);
                        e.appendChild(d);
                        this.thead.insertBefore(e, this.thead.lastChild)
                    }
                    for (b = 0; b < this.thead.children.length; b++) this.thead.children[b].firstChild.style.top = a + _$_669e[57], a += this.thead.children[b].offsetHeight;
                    a = this.thead.offsetHeight - this.thead.lastChild.offsetHeight;
                    for (b = 0; b < this.headers.length; b++) this.headers[b].style.top =
                        a + _$_669e[57]
                }
            };
            h.adjust = function(c, b, a, d) {
                var e, g = this.options.nestedHeaders,
                    f = [],
                    k;
                if (g)
                    if (1 == a)
                        if (d && d.nested)
                            for (e = 0; e < g.length; e++)
                                for (a = 0; a < g[e].length; a++) h.cell.call(this, a, e, {
                                    colspan: d.nested[e][a]
                                });
                        else
                            for (e = 0; e < g.length; e++) {
                                if (a = h.getColumns.call(this, e)) a = a[c], g[e][a] && (d = parseInt(g[e][a].colspan) + b, h.cell.call(this, a, e, {
                                    colspan: d
                                }))
                            } else {
                                for (e = 0; e < g.length; e++)
                                    if (k = h.getColumns.call(this, e)) {
                                        f[e] = [];
                                        for (a = 0; a < b; a++) d = k[c + a], void 0 !== d && (typeof f[e][d] == _$_669e[21] && (f[e][d] = g[e][d].colspan),
                                            g[e][d].colspan--);
                                        d = Object.keys(f[e]);
                                        for (a = 0; a < d.length; a++) h.cell.call(this, d[a], e, {
                                            colspan: g[e][d[a]].colspan
                                        })
                                    } return f
                            }
            };
            return h
        }(),
        fa = function() {
            var h = function(d, e, g, f) {
                    Z(d) ? d = Y.shiftFormula(d, e, g) : f && (d = jSuites.calendar.dateToNum(d), 0 > e || 0 > g ? d-- : d++, d = jSuites.calendar.numToDate(d));
                    return d
                },
                c = function(d, e, g) {
                    if (1 === d.length) {
                        if (1 < e || r.formula && typeof r.formula.TREND !== _$_669e[110]) g ? d[0]-- : d[0]++;
                        return d
                    }
                    e = 1;
                    g = [];
                    for (var f = [], k = 0; k < d.length; k++) f.push(e++);
                    for (k = 0; k < d.length; k++) g.push(e++);
                    return r.formula && typeof r.formula.TREND == _$_669e[110] ? r.formula.TREND(d, f, g) : d
                },
                b = function(d, e, g, f, k) {
                    var l = null,
                        n = [],
                        p = null,
                        q = [];
                    for (l = 0; l < d.length; l++) {
                        if (jSuites.isNumeric(d[l]) && (n.push(d[l]), jSuites.isNumeric(d[l + 1]))) continue;
                        n.length ? (p = c(n, d.length, f), Array.isArray(p) ? (p.forEach(function(v) {
                            q.push(v)
                        }), n = []) : q.push(p)) : (p = void 0 !== k && this.options.columns[k].type === _$_669e[232] ? !0 : !1, f && (e *= -1, g *= -1), p = h(d[l], e * d.length, g * d.length, p), q.push(p))
                    }
                    return q
                },
                a = function(d, e) {
                    var g = function(w) {
                            var x;
                            (x = this.results) ? w = x.indexOf(w): x = this.rows;
                            w--;
                            for (var y = w; 0 <= y; y--)
                                if (w = this.results ? x[y] : y, this.rows[w] && this.rows[w].element) return w;
                            return 0
                        },
                        f = function(w) {
                            var x;
                            (x = this.results) ? w = x.indexOf(w): x = this.rows;
                            w++;
                            for (var y = w; y < x.length; y++)
                                if (w = this.results ? x[y] : y, this.rows[w] && this.rows[w].element) return w;
                            return x[x.length - 1]
                        },
                        k = function(w, x) {
                            for (; 0 <= w; w--)
                                if (this.records[x][w].element && !this.records[x][w].element.classList.contains(_$_669e[65])) return w;
                            return 0
                        },
                        l = function(w, x) {
                            for (; w < this.options.columns.length; w++)
                                if (this.records[x][w].element &&
                                    !this.records[x][w].element.classList.contains(_$_669e[65])) return w;
                            return this.options.columns - 1
                        },
                        n = this.getHighlighted(),
                        p = n[0],
                        q = n[1],
                        v = n[2];
                    n = n[3];
                    d = parseInt(d);
                    e = parseInt(e);
                    null != d && null != e && (0 < d - v ? (l = l.call(this, parseInt(v) + 1, e), k = parseInt(d)) : (l = parseInt(d), k = k.call(this, parseInt(p) - 1, e)), 0 < e - n ? (f = f.call(this, parseInt(n)), g = parseInt(e)) : (f = parseInt(e), g = g.call(this, parseInt(q))), Math.abs(d - v) > Math.abs(e - n) ? (f = parseInt(q), g = parseInt(n)) : (l = parseInt(p), k = parseInt(v)), this.setBorder(l, f, k, g,
                        _$_669e[166]))
                };
            a.execute = function(d) {
                if (!T.call(this.parent, this)) return !1;
                var e = this.getHighlighted(),
                    g = Ka[_$_669e[56]].call(this, e, !1, !1),
                    f = !0,
                    k = !1,
                    l = !1,
                    n = this.borders.cloning;
                if (d)
                    if (n && n.active) {
                        d = n.x1;
                        var p = n.y1,
                            q = n.x2;
                        n = n.y2;
                        e[0] == d ? (k = !0, e[1] > p && (l = !0)) : e[0] > d && (l = !0)
                    } else return this.resetBorders(_$_669e[166], !0), !1;
                else if (n = this.getSelected(), n.length)
                    for (k = !0, d = n[0].x, q = n[n.length - 1].x, p = parseInt(n[n.length - 1].y) + 1, n = parseInt(n[n.length - 1].y); 1 == f;) {
                        if (this.options.data[n + 1]) {
                            this.options.data[n +
                                1][d - 1] || this.options.data[n + 1][q + 1] || (f = !1);
                            for (var v = d; v <= q; v++) this.options.data[n + 1][v] && (f = !1)
                        } else f = !1;
                        f && n++
                    } else return this.resetBorders(_$_669e[166], !0), !1;
                var w = Na.getFromSelection.call(this, e),
                    x = v = 0,
                    y = 0,
                    B = 0,
                    A = [],
                    F = [],
                    u = !0,
                    z = this;
                if (k)
                    for (f = function() {
                            F.push({
                                x: v,
                                y: x,
                                value: A[B],
                                style: w[B][y]
                            });
                            B++;
                            void 0 == g[B] && (B = 0, u && (A = b.call(z, A, 0, 1, l, v)))
                        }, v = d; v <= q; v++) {
                        u = void 0 !== this.options.columns[v].autoIncrement ? this.options.columns[v].autoIncrement : this.parent.config.autoIncrement;
                        B = 0;
                        A = [];
                        for (x = 0; x < g.length; x++) A.push(g[x][y]);
                        l && A.reverse();
                        u && (A = b.call(z, A, 0, 1, l, v));
                        if (l)
                            for (x = n; x >= p; x--) f();
                        else
                            for (x = p; x <= n; x++) f();
                        y++
                    } else
                        for (u = this.parent.config.autoIncrement, f = function() {
                                F.push({
                                    x: v,
                                    y: x,
                                    value: A[y],
                                    style: w[B][y]
                                });
                                y++;
                                void 0 == g[B][y] && (y = 0, u && (A = b(A, 1, 0, l)))
                            }, x = p; x <= n; x++) {
                            y = 0;
                            A = [];
                            for (v = 0; v < g[B].length; v++) A.push(g[B][v]);
                            l && A.reverse();
                            u && (A = b(A, 1, 0, l));
                            if (l)
                                for (v = q; v >= d; v--) f();
                            else
                                for (v = d; v <= q; v++) f();
                            B++
                        }
                F.length && this.setValue(F);
                e[0] < d && (d = e[0]);
                e[1] < p && (p = e[1]);
                e[2] >
                    q && (q = e[2]);
                e[3] > n && (n = e[3]);
                this.resetBorders(_$_669e[166], !0);
                this.updateSelectionFromCoords(d, p, q, n)
            };
            a.end = function(d) {
                fa.event = !1;
                fa.execute.call(this, !0);
                ea.refresh.call(this)
            };
            a.cancel = function(d) {
                fa.event = !1;
                a.reset.call(this, d)
            };
            a.reset = function() {
                this.resetBorders(_$_669e[166], !0)
            };
            return a
        }(),
        Ka = function() {
            var h = function(b) {
                b.getData = h[_$_669e[56]];
                b.getJson = h[_$_669e[56]];
                b.setData = h[_$_669e[11]];
                b.loadData = h.load;
                b.getRowData = h.getRowData;
                b.setRowData = h.setRowData;
                b.getColumnData = h.getColumnData;
                b.setColumnData = h.setColumnData;
                b.refresh = h.refresh;
                b.data = h[_$_669e[56]];
                b.download = h.download;
                b.downloadCSV = h.downloadCSV;
                b.copy = h.copy;
                b.paste = h.paste;
                b.cut = function() {
                    return b.copy(!0)
                }
            };
            h[_$_669e[56]] = function(b, a, d, e) {
                !0 === b && (b = this.getHighlighted());
                var g = b ? !0 : !1,
                    f = [],
                    k = this.options.columns.length,
                    l = this.options.data.length;
                void 0 === e && (e = this.dataType && typeof d == _$_669e[21] ? 1 : 0);
                for (var n = 0; n < l; n++) {
                    var p = null;
                    var q = e ? {} : [];
                    for (var v = 0; v < k; v++) {
                        var w;
                        if (!(w = !g)) {
                            w = b[1];
                            var x = b[2],
                                y = b[3];
                            w = v >= b[0] && v <= x && n >= w && n <= y ? !0 : !1
                        }
                        if (w && (!g || !this.results || -1 != this.results.indexOf(n))) {
                            w = R.call(this, v, n);
                            if (a) {
                                if (g && !1 === this.rows[n].visible) continue;
                                p = qa.processed.call(this, v, n);
                                void 0 !== p && (w = p)
                            }
                            d && w.match && (w.match(d) || w.match(/\n/) || w.match(/"/)) && (w = w.replace(new RegExp(_$_669e[20], _$_669e[14]), _$_669e[233]), w = _$_669e[20] + w + _$_669e[20]);
                            e ? jSuites.path.call(q, _$_669e[3] + ra.call(this, v), w) : q.push(w);
                            p = !0
                        }
                    }
                    p && (d ? f.push(q.join(d)) : f.push(q))
                }
                return d ? f.join(_$_669e[10]) + _$_669e[10] : f
            };
            h.parse =
                function() {
                    for (var b, a = {}, d = 0, e, g = [], f = this.options.data, k = 0; k < f.length; k++) b = null, e = k, f[k].data && (Array.isArray(f[k].data) || typeof f[k].data == _$_669e[109]) && (typeof f[k].id !== _$_669e[21] && (b = parseInt(f[k].id)), typeof f[k].row !== _$_669e[21] && (e = f[k].row), g[e] = f[k].data, f[k] = []), e > d && (d = e), b && (a[e] = b);
                    if (Object.keys(g).length)
                        for (k = 0; k <= d; k++) g[k] && (f[k] = g[k]);
                    d = this.getPrimaryKey();
                    if (!1 !== d)
                        for (d = ra.call(this, d), k = 0; k < f.length; k++) a[k] || !1 === d || (b = parseInt(f[k][d]), 0 < b && (a[k] = b));
                    return a
                };
            h.type =
                function() {
                    for (var b = this.dataType = 0; b < this.options.columns.length; b++) typeof this.options.columns[b].name !== _$_669e[21] && (this.dataType = 1)
                };
            h[_$_669e[11]] = function(b) {
                b = h.standardize(b);
                var a = {},
                    d = [],
                    e = {};
                var g = Object.keys(b[0]).length - this.options.columns.length;
                0 < g && this.insertRow(g);
                var f = b.length - this.options.data.length;
                0 < f && this.insertRow(f);
                for (var k = 0; k < this.options.data.length; k++) {
                    void 0 === b[k] && (b[k] = {
                        row: k,
                        data: []
                    });
                    for (var l = 0; l < this.options.columns.length; l++) void 0 === b[k].data[l] &&
                        (b[k].data[l] = _$_669e[3])
                }
                for (l = 0; l < this.options.columns.length; l++) {
                    var n = ra.call(this, l);
                    e[n] = l
                }
                for (k = 0; k < b.length; k++) {
                    f = b[k].row;
                    n = Object.keys(b[k].data);
                    for (l = 0; l < n.length; l++) g = null, void 0 !== e[n[l]] ? g = e[n[l]] : -1 < n[l] && (g = parseInt(n[l])), null !== g && d.push({
                        x: g,
                        y: f,
                        value: b[k].data[n[l]]
                    });
                    b[k].id && (a[f] = b[k].id)
                }
                za.setId.call(this, a);
                qa[_$_669e[11]].call(this, d, null, !0)
            };
            h.build = function() {
                this.results = null;
                this.rows = [];
                this.records = [];
                this.sequence = 0;
                h.type.call(this);
                if (this.options.data) {
                    var b =
                        h.parse.call(this),
                        a = this.options.data.length;
                    this.options.minDimensions[1] > a && (a = this.options.minDimensions[1]);
                    for (var d = 0; d < a; d++) {
                        var e = b[d] ? b[d] : null;
                        va.row.call(this, d, e)
                    }
                }
            };
            h.load = function(b) {
                this.resetBorders();
                this.resetSelection();
                this.tbody.innerText = _$_669e[3];
                this.options.data = b;
                h.build.call(this);
                J.call(this, !0, !0);
                ha.update.call(this)
            };
            h.refresh = function(b) {
                if (typeof b !== _$_669e[21])
                    for (var a = 0; a < this.options.columns.length; a++) {
                        var d = this.records[b] && this.records[b][a].element ? this.records[b][a].element :
                            null,
                            e = pa[_$_669e[56]].call(this, a, b);
                        typeof e[0].updateCell == _$_669e[110] && e[0].updateCell(d, R.call(this, a, b), a, b, this, e[1]);
                        d && U.applyOverflow.call(this, d, a, b)
                    } else if (this.options.url) {
                        var g = this;
                        1 == this.parent.config.loadingSpin && jSuites.loading.show();
                        jSuites.ajax({
                            url: this.options.url,
                            method: _$_669e[72],
                            dataType: _$_669e[73],
                            success: function(f) {
                                g.setData(f);
                                jSuites.loading.hide()
                            }
                        })
                    }
            };
            h.getRowData = function(b, a) {
                if (typeof this.options.data[b] === _$_669e[21]) return !1;
                if (a) {
                    if (this.dataType) {
                        a = {};
                        for (var d = 0; d < this.options.columns.length; d++) {
                            var e = _$_669e[3] + ra.call(this, d),
                                g = _$_669e[3] + qa.processed.call(this, d, b, !0);
                            jSuites.path.call(a, e, g)
                        }
                    } else
                        for (a = [], d = 0; d < this.options.columns.length; d++) a.push(qa.processed.call(this, d, b, !0));
                    return a
                }
                return JSON.parse(JSON.stringify(this.options.data[b]))
            };
            h.setRowData = function(b, a, d) {
                for (var e = [], g = 0; g < this.options.columns.length; g++) typeof a[g] === _$_669e[21] && (a[g] = _$_669e[3]), e.push({
                    x: g,
                    y: b,
                    value: a[g]
                });
                e.length && this.setValue(e, null, d)
            };
            h.getColumnData =
                function(b, a) {
                    if (typeof this.options.columns[b] === _$_669e[21]) return !1;
                    for (var d = [], e = 0; e < this.options.data.length; e++) {
                        var g = a ? qa.processed.call(this, b, e, !0) : R.call(this, b, e);
                        d.push(g)
                    }
                    return d
                };
            h.setColumnData = function(b, a, d) {
                for (var e = [], g = 0; g < this.rows.length; g++) typeof a[g] === _$_669e[21] && (a[g] = _$_669e[3]), e.push({
                    x: b,
                    y: g,
                    value: a[g]
                });
                e.length && this.setValue(e, null, d)
            };
            h.copy = function(b) {
                var a = this,
                    d = this.getHighlighted(),
                    e = function() {
                        a.setBorder(d[0], d[1], d[2], d[3], _$_669e[115]);
                        a.borders.copying.clear =
                            1 == b ? !0 : !1
                    },
                    g = h[_$_669e[56]].call(this, d, !0, _$_669e[234]),
                    f = G.call(this.parent, _$_669e[235], this, d, g);
                if (f) g = f;
                else if (!1 === f) return !1;
                navigator.clipboard ? navigator.clipboard.writeText(g).then(e) : (this.parent.textarea.value = g, this.parent.textarea.select(), document.execCommand(_$_669e[236]), e());
                return g
            };
            h.paste = function(b, a, d) {
                if (!T.call(this.parent, this)) return !1;
                var e = this,
                    g = function() {
                        e.borders.copying && (e.resetBorders(_$_669e[115], !0), e.borders.copying.clear = !1)
                    },
                    f = null,
                    k = this.options.style,
                    l = null,
                    n = null,
                    p = [];
                if (this.borders.copying) {
                    var q = [this.borders.copying.x1, this.borders.copying.y1, this.borders.copying.x2, this.borders.copying.y2];
                    l = q[0];
                    n = q[1];
                    var v = jSuites.hash(d),
                        w = jSuites.hash(h[_$_669e[56]].call(this, q, !0, _$_669e[234]));
                    if (v != w) g();
                    else {
                        var x = 0,
                            y = 0;
                        d = h[_$_669e[56]].call(this, q, !1, !1);
                        for (v = q[1]; v <= q[3]; v++)
                            if (!this.results || -1 != this.results.indexOf(v)) {
                                p[y] = [];
                                x = 0;
                                for (w = q[0]; w <= q[2]; w++) f = H.getColumnNameFromCoords(w, v), p[y][x] = k && k[f] ? k[f] : _$_669e[237], x++;
                                y++
                            }
                    }
                }
                b = parseInt(b);
                a = parseInt(a);
                if (d) {
                    Array.isArray(d) || (d = H.parseCSV(d, _$_669e[234]));
                    if (v = G.call(this.parent, _$_669e[238], this, d, b, a, p)) d = v;
                    else if (!1 === v) return !1;
                    100 < d.length && !this.options.pagination && !J.limited.call(this) && console.log(_$_669e[239]);
                    if (f = this.getHighlighted()) {
                        x = [];
                        y = [];
                        var B = [];
                        k = [];
                        var A = 0,
                            F = 0;
                        for (v = f[1]; v <= f[3]; v++)
                            if (!this.results || -1 != this.results.indexOf(v)) {
                                x = [];
                                y = [];
                                for (w = f[0]; w <= f[2]; w++) x.push(d[F][A]), p && p[F] && p[F][A] ? y.push(p[F][A]) : y.push(_$_669e[3]), A++, A >= d[0].length && (A = 0);
                                B.push(x);
                                k.push(y);
                                F++;
                                F >= d.length && (F = 0)
                            } 0 == B.length % d.length && 0 == B[0].length % d[0].length && (d = B, p = k)
                    }
                    k = null;
                    f = [];
                    if (this.borders.copying && !0 === this.borders.copying.clear) {
                        for (v = q[1]; v <= q[3]; v++)
                            for (w = q[0]; w <= q[2]; w++)(!this.results || this.results && 0 <= this.results.indexOf(v)) && f.push({
                                x: w,
                                y: v,
                                value: _$_669e[3],
                                style: _$_669e[3]
                            });
                        g()
                    }
                    v = b + d[0].length - this.options.columns.length;
                    0 < v && this.insertColumn(v);
                    v = a + d.length - this.rows.length;
                    0 < v && this.insertRow(v);
                    y = x = b;
                    A = B = a;
                    F = null;
                    this.results && (k = this.results.indexOf(a));
                    for (v = 0; v < d.length; v++)
                        for (w = 0; w < d[v].length; w++) {
                            null !== l && null !== n && d[v][w][0] == _$_669e[18] && (d[v][w] = Y.shiftFormula(d[v][w], b - l, a - n));
                            y = b + w;
                            A = a + v;
                            this.results && (A = this.results[k + v]);
                            var u = {
                                x: y,
                                y: A,
                                value: d[v][w]
                            };
                            p && p[v] && p[v][w] && (u.style = p[v][w]);
                            f.push(u);
                            null === F && ka.isSelected.call(this, y, A, q) && (F = !0)
                        }
                    F && g();
                    f.length && (qa[_$_669e[11]].call(this, f, null, this.parent.config.forceUpdateOnPaste ? !0 : !1), G.call(this.parent, _$_669e[240], this, f), this.updateSelectionFromCoords(x, B, y, A))
                }
            };
            var c = function(b,
                a) {
                if (this.parent.config.allowExport) {
                    void 0 === a && (a = !0);
                    var d = _$_669e[3];
                    if (1 == b || 1 == this.parent.config.includeHeadersOnDownload) d += this.getHeaders().replace(/\s+/gm, _$_669e[161]) + _$_669e[10];
                    d += h[_$_669e[56]].call(this, null, a, this.options.csvDelimiter);
                    a = new Blob([_$_669e[242] + d], {
                        type: _$_669e[243]
                    });
                    window.navigator && window.navigator.msSaveOrOpenBlob ? window.navigator.msSaveOrOpenBlob(a, this.options.csvFileName + _$_669e[244]) : (b = document.createElement(_$_669e[245]), a = URL.createObjectURL(a), b.href =
                        a, b.setAttribute(_$_669e[246], this.options.csvFileName + _$_669e[244]), document.body.appendChild(b), b.click(), b.parentNode.removeChild(b))
                } else console.error(_$_669e[241])
            };
            h.downloadCSV = function(b, a) {
                c.call(this, b, a)
            };
            h.download = function(b, a) {
                r.render ? r.render(this.parent) : c.call(this, b, a)
            };
            h.standardize = function(b) {
                b || (b = []);
                Array.isArray(b) || (b = [b]);
                void 0 == b[0] || Array.isArray(b[0]) || typeof b[0] == _$_669e[109] || (b = [b]);
                for (var a = [], d = 0; d < b.length; d++) {
                    var e = {};
                    Array.isArray(b[d]) || void 0 == b[d].data ?
                        e.data = b[d] : e = b[d];
                    e.row = void 0 == b[d].row ? d : b[d].row;
                    a[e.row] = e
                }
                for (d = 0; d < a.length; d++) b[d] = a[d] || {
                    row: d,
                    data: []
                };
                return b
            };
            return h
        }(),
        ab = function() {
            var h = function(c) {
                c.getLabel = h[_$_669e[56]];
                c.getLabelFromCoords = h[_$_669e[56]]
            };
            h[_$_669e[56]] = function(c, b, a) {
                if (typeof c == _$_669e[25]) {
                    c = H.getCoordsFromColumnName(c);
                    var d = c[0];
                    b = c[1]
                } else d = c;
                return this.getProcessed(d, b, a)
            };
            return h
        }(),
        La = function() {
            var h = function(c) {
                c.getDefinedNames = h[_$_669e[56]];
                c.setDefinedNames = h[_$_669e[11]]
            };
            h[_$_669e[11]] =
                function(c) {
                    if (c.length)
                        if (1 !== this.parent.edition) console.log(_$_669e[247]);
                        else {
                            for (var b = this.parent.config.definedNames, a, d = [], e = 0; e < c.length; e++) {
                                var g = c[e].index.toUpperCase();
                                a = b[g] ? b[g] : null;
                                c[e].value ? (c[e].value = W(c[e].value), this.parent.config.definedNames[g] = c[e].value) : b[g] && delete this.parent.config.definedNames[g];
                                d.push({
                                    index: g,
                                    value: c[e].value,
                                    oldValue: a
                                })
                            }
                            Q.call(this.parent, {
                                worksheet: this,
                                action: _$_669e[102],
                                records: d
                            });
                            L.call(this, _$_669e[102], {
                                data: d
                            });
                            G.call(this.parent, _$_669e[248],
                                this, d)
                        }
                };
            h[_$_669e[56]] = function(c) {
                var b = this.parent.config.definedNames;
                return c ? b[c] : b
            };
            h.updateAll = function(c, b) {
                if (!this.parent.ignoreHistory) {
                    this.parent.historyCascade = !0;
                    var a, d = this.getWorksheetName();
                    if (a = this.parent.config.definedNames) {
                        var e = [],
                            g = Object.keys(a);
                        if (g.length) {
                            for (var f = 0; f < g.length; f++) {
                                var k = Y.update.call(this, a[g[f]], c, b, d);
                                k !== a[g[f]] && (-1 !== k.indexOf(_$_669e[249]) && (k = null), e.push({
                                    index: g[f],
                                    value: k
                                }))
                            }
                            e.length && h[_$_669e[11]].call(this, e)
                        }
                    }
                    this.parent.historyCascade = !1
                }
            };
            h.updateWorksheetName = function(c, b) {
                if (!this.parent.ignoreHistory) {
                    this.parent.historyCascade = !0;
                    var a;
                    if (a = this.parent.config.definedNames) {
                        var d = [],
                            e = Object.keys(a);
                        if (e.length) {
                            for (var g = 0; g < e.length; g++) {
                                var f = Y.updateWorksheetName(a[e[g]], c, b);
                                f !== a[e[g]] && d.push({
                                    index: e[g],
                                    value: f
                                })
                            }
                            d.length && h[_$_669e[11]].call(this, d)
                        }
                    }
                    this.parent.historyCascade = !1
                }
            };
            h.build = function() {
                var c = Object.keys(this.config.definedNames);
                if (c && c.length)
                    if (1 !== this.edition) this.config.definedNames = {}, console.log(_$_669e[247]);
                    else
                        for (var b = 0; b < c.length; b++) {
                            var a = (_$_669e[3] + c[b]).toUpperCase(),
                                d = W(this.config.definedNames[c[b]]);
                            a != c[b] && delete this.config.definedNames[c[b]];
                            this.config.definedNames[a] = d
                        }
            };
            return h
        }(),
        ta = function() {
            var h = [_$_669e[250], _$_669e[251], _$_669e[252], _$_669e[253], _$_669e[254]],
                c = [_$_669e[255], _$_669e[189], _$_669e[256], _$_669e[257], _$_669e[258], _$_669e[259], _$_669e[1], _$_669e[260], _$_669e[37], _$_669e[261], _$_669e[18], _$_669e[15], _$_669e[17], _$_669e[262]],
                b = [],
                a = function() {
                    if (0 < b.length) {
                        for (var f,
                                k = 0; k < b.length; k++) f = -1 === b[k].indexOf(_$_669e[19]) ? this : I(b[k].split(_$_669e[19])[0]), ca.destroy.call(f, b[k]);
                        b = []
                    }
                },
                d = function(f) {
                    a.call(this);
                    for (var k = 0, l = [], n, p, q, v = 0; v < f.children.length; v++)
                        if (n = f.children[v].getAttribute(_$_669e[263])) {
                            l[n] || (l[n] = h[k++], 4 < k && (k = 0));
                            f.children[v].style.color = l[n]; - 1 === n.indexOf(_$_669e[19]) ? (p = this, q = n) : (p = n.split(_$_669e[19]), q = p[1], p = I(p[0]));
                            if (0 <= q.indexOf(_$_669e[164])) {
                                var w = q.split(_$_669e[164]);
                                q = H.getCoordsFromColumnName(w[0]);
                                w = H.getCoordsFromColumnName(w[1])
                            } else w =
                                q = H.getCoordsFromColumnName(q);
                            ka[_$_669e[11]].call(p, q[0], q[1], w[0], w[1], null, n, l[n]);
                            b.push(n)
                        }
                },
                e = function(f) {
                    this.worksheet && (a.call(this.worksheet), typeof this.onchange == _$_669e[110] && this.onchange.call(this, this.innerText, f), this.worksheet = null)
                },
                g = function(f, k) {
                    k && typeof k != _$_669e[110] || (k = {
                        type: _$_669e[264],
                        onchange: k
                    });
                    k.type === _$_669e[264] ? (f.classList.add(_$_669e[265]), f.addEventListener(_$_669e[266], function(l) {
                        l = l.changedTouches && l.changedTouches[0] ? l.changedTouches[0].clientX : l.clientX;
                        var n = this.getBoundingClientRect();
                        24 > n.width - (l - n.left) && (g.current = f, f.classList.add(_$_669e[267]), f.classList.add(_$_669e[121]))
                    })) : (f.addEventListener(_$_669e[268], function(l) {
                        Z(f.innerText) && (f.classList.add(_$_669e[121]), r.current && setTimeout(function() {
                            f.worksheet = r.current;
                            g.parse.call(r.current, f);
                            var n = document.createTextNode(_$_669e[3]);
                            f.appendChild(n);
                            jSuites.focus(f)
                        }))
                    }), f.addEventListener(_$_669e[269], function(l) {
                        Z(f.innerText) ? f.classList.add(_$_669e[121]) : f.classList.contains(_$_669e[121]) &&
                            (f.classList.remove(_$_669e[121]), f.worksheet && a.call(f.worksheet));
                        l.preventDefault()
                    }), f.addEventListener(_$_669e[116], function(l) {
                        27 == l.which && (f.innerText = _$_669e[3], e.call(f, !1))
                    }));
                    f.classList.add(_$_669e[270]);
                    f.classList.add(_$_669e[269]);
                    f.setAttribute(_$_669e[271], !0);
                    f.addEventListener(_$_669e[272], function(l) {
                        e.call(f, !0)
                    });
                    f.val = function(l) {
                        if (void 0 === l) return f.innerText;
                        f.innerText = l
                    };
                    k.onchange && (f.onchange = k.onchange);
                    k.onupdate && (f.onupdate = k.onupdate)
                };
            g.current = null;
            g.parse = function(f) {
                for (var k =
                        _$_669e[3], l = _$_669e[3], n = window.getSelection(), p = document.createRange(), q = H.getCaretIndex.call(this.parent, f), v = Y.tokenize(f.innerText.replace(/(\r\n|\n|\r)/gm, _$_669e[3])), w = 0; w < v.length; w++) v[w] && Y.tokenIdentifier.test(v[w].trim()) ? (w == v.length - 1 && (l = _$_669e[273]), k += _$_669e[274] + l + _$_669e[275] + v[w].trim().replace(/\$/g, _$_669e[3]) + _$_669e[276] + v[w] + _$_669e[277]) : k += v[w];
                f.innerHTML = k;
                k = null;
                for (w = 0; w < f.childNodes.length && !(k = f.childNodes[w], l = k.tagName ? k.innerText.length : k.length, 0 >= q - l); w++) q -=
                    l;
                k && (k.tagName ? p.setStart(k.firstChild, q) : (q > k.length && (q = k.length), p.setStart(k, q)), n.removeAllRanges(), n.addRange(p));
                d.call(this, f)
            };
            g.close = function(f) {
                a.call(this);
                f ? f.classList.remove(_$_669e[121]) : g.current && (typeof g.current.onchange == _$_669e[110] && g.current.onchange.call(g.current), g.current.classList.remove(_$_669e[267]), g.current.classList.remove(_$_669e[121]), g.current.worksheet = null, g.current = null)
            };
            g.getName = function(f, k, l, n, p) {
                f = H.getColumnNameFromCoords(f, k);
                l = H.getColumnNameFromCoords(l,
                    n);
                p.ctrlKey && (f = l);
                return f !== l ? (p = f + _$_669e[164] + l, p = Y.getTokensFromRange.call(this, p), Y.getRangeFromTokens.call(this, p)) : f
            };
            g.range = function(f, k, l, n, p) {
                g.current.getAttribute(_$_669e[278]) == _$_669e[279] && (f = l, k = n);
                f = g.getName.call(this, f, k, l, n, p);
                k = document.createElement(_$_669e[280]);
                k.innerText = f;
                k.setAttribute(_$_669e[263], f);
                g.current.innerText = _$_669e[3];
                g.current.appendChild(k);
                d.call(this, g.current);
                jSuites.focus(k);
                p.preventDefault()
            };
            g.update = function(f, k, l, n, p, q, v) {
                q.getAttribute(_$_669e[278]) ==
                    _$_669e[279] && (f = l, k = n);
                var w = jSuites.getNode();
                if (w) {
                    var x = g.getName.call(this, f, k, l, n, p);
                    this !== v && (x = v.getWorksheetName() + _$_669e[19] + x);
                    f = function(y) {
                        var B = document.createElement(_$_669e[280]);
                        B.innerText = x;
                        B.setAttribute(_$_669e[263], x);
                        w.parentNode === q && !1 === y ? w.parentNode.insertBefore(B, w.nextSibling) : q.appendChild(B);
                        w = B
                    };
                    if (w.getAttribute(_$_669e[263])) p.ctrlKey ? (q.appendChild(document.createTextNode(_$_669e[37])), f(!0)) : (w.innerText = x, w.setAttribute(_$_669e[263], x));
                    else {
                        if (!w.innerText ||
                            w.innerText !== _$_669e[3] && -1 == c.indexOf(w.innerText.slice(-1)) || w.nextElementSibling && w.nextElementSibling.getAttribute(_$_669e[263])) return !1;
                        f(!1)
                    }
                    jSuites.focus(w);
                    d.call(this, q);
                    p.preventDefault();
                    typeof q.onupdate == _$_669e[110] && q.onupdate.call(q, q.innerText);
                    return !0
                }
                return !1
            };
            g.onkeydown = function(f) {
                f.target.worksheet = this;
                if (f.ctrlKey) {
                    if (67 == f.which || 88 == f.which) f = window.getSelection().toString().replace(/(\r\n|\n|\r)/gm, _$_669e[3]), navigator.clipboard ? navigator.clipboard.writeText(f) : (this.parent.textarea.value =
                        f, this.parent.textarea.select(), document.execCommand(_$_669e[236]))
                } else if (36 == f.which) {
                    var k = window.getSelection(),
                        l = document.createRange();
                    l.setStart(f.target, 0);
                    k.removeAllRanges();
                    k.addRange(l)
                } else 35 == f.which ? jSuites.focus(f.target) : 187 == f.which ? !f.target.innerText.replace(/(\r\n|\n|\r)/gm, _$_669e[3]) && f.target.getAttribute(_$_669e[123]) && f.target.removeAttribute(_$_669e[123]) : 13 == f.which && f.target.blur()
            };
            g.palette = function(f) {
                h = f
            };
            return g
        }(),
        Y = function() {
            var h = new RegExp(/^(('.*?'!)|(\w*!))?(\$?[A-Z]+\$?[0-9]+)(:\$?[A-Z]+\$?[0-9]+)?$/i),
                c = function(f) {
                    f = W(f);
                    for (var k = _$_669e[3], l = [], n = 0, p = 0; p < f.length; p++) {
                        if (f[p] == _$_669e[20] || f[p] == _$_669e[162]) n = 0 == n ? 1 : 0;
                        var q;
                        if (q = 0 == n) q = f[p], (q = q == _$_669e[255] || q == _$_669e[189] || q == _$_669e[256] || q == _$_669e[257] || q == _$_669e[258] || q == _$_669e[12]) || (q = f[p], q = q == _$_669e[259] || q == _$_669e[281] || q == _$_669e[1] || q == _$_669e[2] || q == _$_669e[260] || q == _$_669e[282] || q == _$_669e[37] || q == _$_669e[261] || q == _$_669e[18] || q == _$_669e[15] || q == _$_669e[17] || q == _$_669e[262]);
                        q ? (k && (l.push(k), k = _$_669e[3]), l.push(f[p])) :
                            k += f[p]
                    }
                    k && l.push(k);
                    for (p = 0; p < l.length; p++) 0 == n && (l[p] = l[p].trim());
                    return l
                },
                b = function(f) {
                    f.setFormula = b[_$_669e[11]];
                    f.updateFormula = b.update;
                    f.executeFormula = b.execute;
                    f.resetArray = b.resetArray;
                    f.formula = [];
                    f.tracking = [];
                    f.setTracking = b.setTracking
                };
            b.run = function(f, k) {
                var l = _$_669e[3],
                    n = Object.keys(k);
                if (n.length) {
                    for (var p = {}, q = 0; q < n.length; q++) {
                        var v = n[q].replace(/!/g, _$_669e[39]);
                        0 < v.indexOf(_$_669e[39]) && (v = v.split(_$_669e[39]), p[v[0]] = !0)
                    }
                    v = Object.keys(p);
                    for (q = 0; q < v.length; q++) l += _$_669e[283] +
                        v[q] + _$_669e[284];
                    for (q = 0; q < n.length; q++) v = n[q].replace(/!/g, _$_669e[39]), null === k[n[q]] || jSuites.isNumeric(k[n[q]]) || (p = k[n[q]].match(/(('.*?'!)|(\w*!))?(\$?[A-Z]+\$?[0-9]*):(\$?[A-Z]+\$?[0-9]*)?/g)) && p.length && (k[n[q]] = d(p, k[n[q]])), l = 0 < v.indexOf(_$_669e[39]) ? l + (v + _$_669e[285] + k[n[q]] + _$_669e[286]) : l + (_$_669e[283] + v + _$_669e[285] + k[n[q]] + _$_669e[286])
                }
                f = f.replace(/\$/g, _$_669e[3]);
                (p = f.match(/(('.*?'!)|(\w*!))?(\$?[A-Z]+\$?[0-9]*):(\$?[A-Z]+\$?[0-9]*)?/g)) && p.length && (f = d(p, f));
                f = f.replace(/!/g,
                    _$_669e[39]);
                f = (new Function(l + _$_669e[287] + f))();
                null === f && (f = 0);
                return f
            };
            b[_$_669e[11]] = function() {};
            var a = function(f) {
                    f = W(f, !0);
                    return f.replace(/\$/g, _$_669e[3])
                },
                d = function(f, k) {
                    for (var l = 0; l < f.length; l++) {
                        var n = b.getTokensFromRange.call(this, f[l]);
                        k = k.replace(f[l], _$_669e[260] + n.join(_$_669e[37]) + _$_669e[282])
                    }
                    return k
                },
                e = function(f, k, l) {
                    var n = this.config.definedNames;
                    if (n && 0 < Object.keys(n).length) {
                        for (var p = 0; p < f.length; p++) f[p] && n && n[f[p]] && (k = k.replace(f[p], n[f[p]]), l && (l[f[p]] = n[f[p]]));
                        return a(k)
                    }
                    return k
                };
            b.resetArray = function(f, k) {
                var l;
                if (l = this.records[k][f].a) {
                    for (var n = 0; n < l.length; n++)
                        for (var p = 0; p < l[n].length; p++) this.options.data[k + n][f + p] || (this.records[k + n][f + p].v = _$_669e[3], this.records[k + n][f + p].element && (this.records[k + n][f + p].element.innerText = _$_669e[3]));
                    this.records[k][f].a = null
                }
            };
            b.execute = function(f, k, l, n, p) {
                if (1 == this.parent.processing) return null !== k && null !== l && this.parent.queue.push([this, k, l]), _$_669e[3];
                if (typeof this.parent.config.onbeforeformula == _$_669e[110]) {
                    var q =
                        G.call(this.parent, _$_669e[288], this, f, k, l);
                    if (!1 === q) return f;
                    q && (f = q)
                }
                var v = [],
                    w = function(E, K, N, V) {
                        var O = null,
                            oa = null;
                        if (void 0 != K && void 0 != N) {
                            O = H.getColumnNameFromCoords(K, N);
                            oa = V.getWorksheetName() + _$_669e[39] + O;
                            if (v[oa]) throw {
                                error: _$_669e[289],
                                message: _$_669e[290],
                                reference: oa
                            };
                            v[oa] = !0
                        }
                        p || (E = a(E));
                        var X = E,
                            S = X.match(/[A-Z_]+[A-Z0-9_\.!]*[(]*/g);
                        S && S.length && (X = e.call(V.parent, S, X, x));
                        (S = X.match(/(('.*?'!)|(\w*!))?(\$?[A-Z]+\$?[0-9]*):(\$?[A-Z]+\$?[0-9]*)?/g)) && S.length && (X = d.call(V, S, X));
                        S = [];
                        X = c(X);
                        for (var ba = 0; ba < X.length; ba++) X[ba] && h.test(X[ba]) && S.push(X[ba]);
                        if (S)
                            for (X = 0; X < S.length; X++) {
                                var da = S[X].split(_$_669e[19]);
                                da[1] ? (K = da[1], ba = I(da[0]), N = V.getWorksheetName() + _$_669e[39] + O, da = S[X]) : (K = da[0], ba = V, N = O, da = V.getWorksheetName() + _$_669e[39] + da[0]);
                                if (da === oa) throw {
                                    error: _$_669e[291],
                                    message: _$_669e[292]
                                };
                                if (typeof ba == _$_669e[21] || typeof ba.formula == _$_669e[21]) throw {
                                    error: _$_669e[293],
                                    message: _$_669e[294]
                                };
                                N && (ba.formula[K] || (ba.formula[K] = []), 0 > ba.formula[K].indexOf(N) && ba.formula[K].push(N));
                                x[S[X]] || (N = H.getCoordsFromColumnName(K), K = N[0], N = N[1], ba.rows.length > N && ba.options.columns.length > K && (da = ba.records[N] && void 0 !== ba.records[N][K].v ? ba.records[N][K].v : R.call(ba, K, N), Z(da) ? x[S[X]] = w(da, K, N, ba) : (da === _$_669e[3] || null === da || void 0 === da ? da = null : jSuites.isNumeric(da) || (da = _$_669e[20] + da + _$_669e[20], y[S[X]] = _$_669e[69]), x[S[X]] = da)))
                            }
                        return E.substr(1)
                    },
                    x = {},
                    y = {};
                try {
                    f = w(f, k, l, this);
                    if (f == _$_669e[3]) return _$_669e[3];
                    0 == f.indexOf(_$_669e[295]) && (p = !0);
                    var B = p ? b.run(f, x) : r.formula(f, x,
                        k, l, this, y);
                    if (!1 !== n && void 0 != k && void 0 != l)
                        if (B instanceof Date) this.records[l][k].v = jSuites.calendar.dateToNum(B);
                        else if (Array.isArray(B)) {
                        n = [];
                        q = null;
                        for (var A = 0; A < B.length; A++)
                            for (var F = 0; F < B[A].length; F++) {
                                (0 < F || 0 < A) && this.options.data[l + A][k + F] !== _$_669e[3] && (q = _$_669e[296]);
                                var u = H.getColumnNameFromCoords(k + F, l + A);
                                n.push(u)
                            }
                        if (!q) {
                            u = H.getColumnNameFromCoords(k, l);
                            var z = Object.keys(this.formula);
                            for (A = 0; A < z.length; A++) {
                                var C = this.formula[z[A]].indexOf(u);
                                if (0 <= C)
                                    for (F = 0; F < n.length; F++) C = this.formula[z[A]].indexOf(n[F]),
                                        -1 == C && this.formula[z[A]].push(n[F])
                            }
                            for (A = 0; A < B.length; A++)
                                for (F = 0; F < B[A].length; F++) this.records[l + A][k + F].v = B[A][F], this.records[l + A][k + F].element && (this.records[l + A][k + F].element.innerText = B[A][F], jSuites.isNumeric(B[A][F]) ? this.records[l + A][k + F].element.classList.add(_$_669e[297]) : this.records[l + A][k + F].element.classList.remove(_$_669e[297]));
                            this.records[l][k].a = B;
                            q = B[0][0];
                            this.records[l][k].v = q
                        }
                        B = q
                    } else this.records[l][k].v = B;
                    return B
                } catch (E) {
                    return 1 == this.parent.config.debugFormulas && console.log(f,
                        x, E), E && E.error ? E.error : _$_669e[293]
                }
            };
            b.update = function(f, k, l, n) {
                var p = null,
                    q = null,
                    v = this.getWorksheetName(),
                    w = function(F) {
                        l && l[F] && (F = _$_669e[249]);
                        k[F] && (F = k[F]);
                        return F
                    },
                    x = function(F, u) {
                        if (-1 === u.indexOf(_$_669e[298])) return F;
                        u = u.split(_$_669e[298]);
                        if (3 === u.length) {
                            u = _$_669e[298];
                            var z = _$_669e[298]
                        } else u[0] === _$_669e[3] ? (u = _$_669e[298], z = _$_669e[3]) : (u = _$_669e[3], z = _$_669e[298]);
                        return u + F.match(/[A-Z]+/g) + z + F.match(/[0-9]+/g)
                    };
                f = W(f, !0);
                f = c(f);
                for (var y = 0; y < f.length; y++)
                    if (f[y] && h.test(f[y])) {
                        if (-1 ==
                            f[y].indexOf(_$_669e[19])) {
                            if (q = f[y], p = _$_669e[3], n && n != v) continue
                        } else if (q = f[y].split(_$_669e[19]), p = q[0], q = q[1], n && n != p) continue;
                        var B = q;
                        if (0 <= q.indexOf(_$_669e[164])) {
                            q = b.getTokensFromRange.call(this, q);
                            for (var A = 0; A < q.length; A++) q[A] = w(q[A]);
                            q = b.getRangeFromTokens.call(this, q);
                            q !== _$_669e[249] && -1 !== B.indexOf(_$_669e[298]) && (B = B.split(_$_669e[164]), q = q.split(_$_669e[164]), q = x(q[0], B[0]) + _$_669e[164] + x(q[1], B[1]))
                        } else - 1 !== B.indexOf(_$_669e[298]) ? (q = w(q.replace(/\$/g, _$_669e[3])), q = x(q, B)) : q =
                            w(q);
                        p && (q = p + _$_669e[19] + q);
                        f[y] = q
                    } return f.join(_$_669e[3])
            };
            b.getChain = function(f) {
                if (void 0 === f) return [];
                var k = [],
                    l = [],
                    n = [],
                    p = function(w) {
                        if (0 <= w.indexOf(_$_669e[39])) {
                            var x = w.split(_$_669e[39]),
                                y = x.pop(),
                                B = x.join(_$_669e[39]);
                            x = I(B)
                        } else x = this, B = x.getWorksheetName(), y = w, w = B + _$_669e[39] + w;
                        if (x.formula[y] && !n[w]) {
                            var A = !1;
                            n[w] = !0;
                            for (w = 0; w < x.formula[y].length; w++) {
                                var F = x.formula[y][w]; - 1 == F.indexOf(_$_669e[39]) && (F = B + _$_669e[39] + F);
                                if (!l[F]) {
                                    var u = x.getValue(F);
                                    Z(u) ? k.push([F, u, b.getTokens.call(x,
                                        u, B)]) : (A = !0, x.formula[y][w] = null);
                                    p.call(x, F);
                                    l[F] = !0
                                }
                            }
                            1 == A && (x.formula[y] = x.formula[y].filter(function(z) {
                                return null != z
                            }))
                        }
                    },
                    q = Object.keys(f);
                for (f = 0; f < q.length; f++) p.call(this, q[f]);
                q = 0;
                for (f = k.length - 1; 0 <= f; f--) {
                    for (var v = 0; v < f; v++)
                        if (0 <= k[v][2].indexOf(k[f][0])) {
                            k.splice(v, 0, k.splice(f, 1)[0]);
                            f = k.length;
                            break
                        } q++;
                    if (5E4 < q) {
                        console.error(_$_669e[299]);
                        break
                    }
                }
                l = [];
                for (f = 0; f < k.length; f++) l[k[f][0]] = k[f][1];
                return l
            };
            b.updateAll = function(f, k) {
                var l, n, p, q, v, w = {},
                    x = this.getWorksheetName();
                for (v = 0; v <
                    r.spreadsheet.length; v++) {
                    var y = r.spreadsheet[v].worksheets;
                    for (p = 0; p < y.length; p++) {
                        var B = y[p].getWorksheetName();
                        for (n = 0; n < y[p].options.data.length; n++)
                            for (l = 0; l < y[p].options.data[n].length; l++)
                                if ((q = y[p].options.data[n][l]) && Z(q)) {
                                    var A = b.update.call(y[p], q, f, k, x);
                                    A != q && (w[B] || (w[B] = []), w[B].push({
                                        x: l,
                                        y: n,
                                        value: A,
                                        force: !0
                                    }))
                                }
                    }
                }
                y = Object.keys(w);
                for (l = 0; l < y.length; l++)
                    if (w[y[l]].length) {
                        p = I(y[l]);
                        for (n = 0; n < w[y[l]].length; n++) R.call(p, w[y[l]][n].x, w[y[l]][n].y, w[y[l]][n].value);
                        L.call(p, _$_669e[136], {
                            data: w[y[l]]
                        })
                    } w = [];
                l = Object.keys(this.formula);
                y = [];
                n = [];
                for (p = 0; p < l.length; p++)
                    if (A = l[p], q = this.formula[A], k && k[A])
                        for (v = 0; v < q.length; v++) k && k[q[v]] || (f[q[v]] && (q[v] = f[q[v]]), n[q[v]] = !0);
                    else {
                        B = [];
                        for (v = 0; v < q.length; v++) k && k[q[v]] || (f[q[v]] && (q[v] = f[q[v]]), n[q[v]] = !0, B.push(q[v]));
                        B.length && (f[A] && (A = f[A]), y[A] = B)
                    } this.formula = y;
                B = this.getWorksheetName();
                for (var F = 0; F < r.spreadsheet.length; F++)
                    for (var u = r.spreadsheet[F].worksheets, z = 0; z < u.length; z++)
                        if (u[z].getWorksheetName() !== B)
                            for (l = Object.keys(u[z].formula),
                                p = 0; p < l.length; p++) {
                                y = [];
                                q = u[z].formula[l[p]];
                                for (v = 0; v < q.length; v++) A = q[v].split(_$_669e[39]), A[0] == B && A[1] ? k && k[A[1]] || (f[A[1]] && (A[1] = f[A[1]]), y.push(A[0] + _$_669e[39] + A[1])) : y.push(q[v]);
                                u[z].formula[l[p]] = y
                            }
                y = Object.keys(n);
                for (l = 0; l < y.length; l++) A = y[l].split(_$_669e[39]), A[1] ? (q = A[1], f = I(A[0]), k = A[0]) : (q = A[0], f = this, k = x), A = H.getCoordsFromColumnName(q), n = A[0], p = A[1], q = f.options.data[p][n], A = Y.execute.call(f, q, n, p, !1), A !== f.records[p][n].v && (w[k] || (w[k] = []), w[k].push({
                    x: n,
                    y: p,
                    value: q,
                    force: !0
                }));
                y = Object.keys(w);
                if (y.length)
                    for (l = 0; l < y.length; l++) p = I(y[l]), w[y[l]] = qa.applyValues.call(p, w[y[l]]), qa.setValueChained.call(p, w[y[l]])
            };
            b.updateWorksheetName = function(f, k, l) {
                for (var n = c(f), p = 0; p < n.length; p++) f = n[p].split(_$_669e[19]), f[1] && (f[0] = f[0].replace(new RegExp(_$_669e[162], _$_669e[14]), _$_669e[3]), f[0].toUpperCase() == k.toUpperCase() && (f[0] = l, 0 <= f[0].indexOf(_$_669e[161]) && (f[0] = _$_669e[162] + f[0] + _$_669e[162]), n[p] = f.join(_$_669e[19])));
                return n.join(_$_669e[3])
            };
            b.updateWorksheetNames = function(f,
                k) {
                var l, n;
                for (n = 0; n < r.spreadsheet.length; n++) {
                    var p = r.spreadsheet[n].worksheets;
                    for (var q = 0; q < p.length; q++) {
                        var v = [];
                        for (var w = 0; w < p[q].options.data.length; w++)
                            for (var x = 0; x < p[q].options.data[w].length; x++)
                                if ((l = p[q].options.data[w][x]) && Z(l)) {
                                    var y = b.updateWorksheetName(l, f, k);
                                    y != l && (v.push({
                                        x,
                                        y: w,
                                        value: y,
                                        force: !0
                                    }), R.call(p[q], x, w, y))
                                } v.length && L.call(p[q], _$_669e[136], {
                            data: v
                        })
                    }
                }
            };
            var g = function(f, k, l) {
                var n = _$_669e[3],
                    p = _$_669e[3],
                    q = f.split(_$_669e[298]);
                if (3 === q.length) return f;
                2 === q.length &&
                    (q[0] ? (l = 0, p = _$_669e[298]) : (k = 0, n = _$_669e[298]), f = q.join(_$_669e[3]));
                f = H.getCoordsFromColumnName(f);
                k = f[0] + k;
                f = f[1] + l;
                return f = 0 > k || 0 > f ? _$_669e[249] : n + H.getColumnName(k) + p + (f + 1)
            };
            b.shiftFormula = function(f, k, l) {
                f = c(f, !0);
                for (var n, p, q = 0; q < f.length; q++) h.test(f[q]) && (p = -1 == f[q].indexOf(_$_669e[19]) ? [_$_669e[3], f[q]] : f[q].split(_$_669e[19]), -1 == p[1].indexOf(_$_669e[164]) ? p[1] = g(p[1], k, l) : (n = p[1].split(_$_669e[164]), p[1] = g(n[0], k, l) + _$_669e[164] + g(n[1], k, l)), f[q] = p[0] ? p.join(_$_669e[19]) : p[1]);
                return f.join(_$_669e[3])
            };
            b.getTokensFromRange = function(f) {
                var k = this,
                    l = _$_669e[3],
                    n = _$_669e[3];
                0 <= f.indexOf(_$_669e[19]) ? n = _$_669e[19] : 0 <= f.indexOf(_$_669e[39]) && (n = _$_669e[39]);
                f = f.replace(/\$/g, _$_669e[3]);
                n && (l = f.split(n), f = l.pop(), l = l.join(_$_669e[39]), k = I(l), l += n);
                if (-1 == f.indexOf(_$_669e[164])) return [l + f];
                n = [];
                f = f.split(_$_669e[164]);
                var p = H.getCoordsFromColumnName(f[0]),
                    q = H.getCoordsFromColumnName(f[1]);
                if (p[0] <= q[0]) {
                    f = p[0];
                    var v = q[0]
                } else f = q[0], v = p[0];
                if (null === p[1] && null == q[1]) {
                    var w = 0;
                    p = 0;
                    k.options && k.options.data &&
                        (p = k.options.data.length - 1)
                } else p[1] <= q[1] ? (w = p[1], p = q[1]) : (w = q[1], p = p[1]);
                for (k = w; k <= p; k++)
                    for (w = f; w <= v; w++) n.push(l + H.getColumnNameFromCoords(w, k));
                return n
            };
            b.getRangeFromTokens = function(f) {
                f = f.filter(function(n) {
                    return n != _$_669e[249]
                });
                for (var k = _$_669e[3], l = 0; l < f.length; l++) 0 <= f[l].indexOf(_$_669e[19]) ? (k = f[l].split(_$_669e[19]), f[l] = k.pop(), k = k[0] + _$_669e[19]) : 0 <= f[l].indexOf(_$_669e[39]) && (k = f[l].split(_$_669e[39]), f[l] = k.pop(), k = k.join(_$_669e[39]) + _$_669e[39]);
                f.sort(function(n, p) {
                    n = H.getCoordsFromColumnName(n);
                    p = H.getCoordsFromColumnName(p);
                    return n[1] > p[1] ? 1 : n[1] < p[1] ? -1 : n[0] > p[0] ? 1 : n[0] < p[0] ? -1 : 0
                });
                return f.length ? k + (f[0] + _$_669e[164] + f[f.length - 1]) : _$_669e[249]
            };
            b.getTokens = function(f, k) {
                var l = a(f);
                f.replace(/!/g, _$_669e[39]);
                (f = l.match(/[A-Z_]+[A-Z0-9_\.]*/g)) && f.length && (l = e.call(this.parent, f, l));
                (f = l.match(/('.*?'?|\w*\.)?(\$?[A-Z]+\$?[0-9]*):(\$?[A-Z]+\$?[0-9]*)?/g)) && f.length && (l = d.call(this, f, l));
                if (f = l.match(/('.*?'?|\w*\.)?(\$?[A-Z]+\$?[0-9]+)(:\$?[A-Z]+\$?[0-9]+)?/g))
                    for (l = 0; l < f.length; l++) - 1 ==
                        f[l].indexOf(_$_669e[39]) && k && (f[l] = k + _$_669e[39] + f[l]);
                return f
            };
            b.setTracking = function() {
                var f = H.getColumnNameFromCoords(this.x, this.y); - 1 === this.instance.tracking.indexOf(f) && this.instance.tracking.push(f)
            };
            b.tracking = function(f, k) {
                for (var l = this.tracking, n = [], p = 0; p < l.length; p++) k && k[l[p]] || (f[l[p]] ? n.push(f[l[p]]) : n.push(l[p]));
                this.tracking = n;
                b.evaluate.call(this)
            };
            b.evaluate = function() {
                var f = this.tracking;
                if (f)
                    for (var k = 0; k < f.length; k++) {
                        var l = this.getValue(f[k]);
                        l && Z(l) ? this.setValue(f[k],
                            l) : f.splice(k, 1)
                    }
            };
            b.tokenIdentifier = h;
            b.tokenize = c;
            return b
        }(),
        qa = function() {
            var h = function(c) {
                c.name = ra;
                c.value = R;
                c.getValue = h[_$_669e[56]];
                c.getValueFromCoords = h.getFromCoords;
                c.setValue = h[_$_669e[11]];
                c.setValueFromCoords = h.setFromCoords;
                c.setCheckRadioValue = h.setCheckRadio;
                c.getProcessed = h.processed
            };
            h[_$_669e[56]] = function(c, b) {
                var a = this;
                if (typeof c == _$_669e[109]) {
                    var d = c.x;
                    c = c.y
                } else 0 <= c.indexOf(_$_669e[39]) && (a = c.split(_$_669e[39]), c = a.pop(), a = I(a.join(_$_669e[39]))), c = H.getCoordsFromColumnName(c),
                    d = c[0], c = c[1];
                return h.getFromCoords.call(a, d, c, b)
            };
            h.getFromCoords = function(c, b, a) {
                var d = null;
                null != c && null != b && (d = a ? h.processed.call(this, c, b) : R.call(this, c, b));
                return d
            };
            h[_$_669e[11]] = function(c, b, a) {
                if (!T.call(this.parent, this)) return !1;
                var d = null,
                    e = null,
                    g = null,
                    f = [];
                a = a ? !0 : !1;
                var k = function(p, q) {
                    var v = null;
                    typeof p == _$_669e[25] ? (g = H.getCoordsFromColumnName(p), d = g[0], e = g[1]) : p.tagName ? (d = p.getAttribute(_$_669e[63]), e = p.getAttribute(_$_669e[59])) : p.element && p.element.tagName ? (d = p.element.getAttribute(_$_669e[63]),
                        e = p.element.getAttribute(_$_669e[59])) : (d = p.x, e = p.y, void 0 !== p.value && (q = p.value), void 0 !== p.style && (v = p.style));
                    p = {
                        x: d,
                        y: e,
                        value: q,
                        force: a
                    };
                    null !== v && (p.style = v);
                    f.push(p)
                };
                if (c && Array.isArray(c))
                    for (var l = 0; l < c.length; l++) k(c[l], b);
                else if (typeof c == _$_669e[109] && void 0 == c.x && void 0 == c.tagName)
                    for (b = Object.keys(c), l = 0; l < b.length; l++) k(b[l], c[b[l]]);
                else k(c, b);
                if (f.length) {
                    f = h.applyValues.call(this, f);
                    this.parent.config.parseFormulas && h.setValueChained.call(this, f);
                    this.refreshBorders();
                    var n = this;
                    setTimeout(function() {
                        ca.refresh.call(n)
                    }, 250);
                    if (f.length) {
                        Q.call(this.parent, {
                            worksheet: n,
                            action: _$_669e[97],
                            records: f,
                            selection: this.selectedCell
                        });
                        c = this.getPrimaryKey();
                        if (!1 !== c)
                            for (l = 0; l < f.length; l++) f[l].x === c && f[l].value && P.call(this, f[l]);
                        L.call(this, _$_669e[97], {
                            data: f
                        });
                        ha.update.call(this, !0);
                        G.call(this.parent, _$_669e[300], this, f)
                    }
                }
            };
            h.setFromCoords = function(c, b, a, d) {
                var e = [];
                e.push({
                    x: c,
                    y: b,
                    value: a
                });
                h[_$_669e[11]].call(this, e, null, d)
            };
            h.setCheckRadio = function() {
                for (var c = [], b = this.getSelected(),
                        a = 0; a < b.length; a++) {
                    var d = b[a].x,
                        e = b[a].y;
                    if (!this.options.columns[d].readOnly && this.options.columns[d].type == _$_669e[301]) {
                        var g = R.call(this, d, e) ? !1 : !0;
                        c.push({
                            x: d,
                            y: e,
                            value: g
                        })
                    }
                }
                c.length && h[_$_669e[11]].call(this, c)
            };
            h.applyValues = function(c) {
                for (var b, a, d = this.options.style || {}, e = [], g = {}, f = 0; f < c.length; f++)
                    if (b = c[f], a = this.updateCell(b.x, b.y, b.value, b.force), void 0 !== a.value) {
                        if (void 0 !== b.style) {
                            var k = H.getColumnNameFromCoords(b.x, b.y);
                            a.style = g[k] = b.style;
                            a.oldStyle = d && d[k] ? d[k] : _$_669e[237]
                        }
                        e.push(a)
                    } 0 <
                    Object.keys(g).length && Na.update.call(this, g);
                return e
            };
            h.processed = function(c, b, a) {
                var d = R.call(this, c, b);
                if (null === d) return null;
                var e = pa[_$_669e[56]].call(this, c, b),
                    g = e[1];
                if (g.type != _$_669e[119] && g.type != _$_669e[302] && g.type != _$_669e[232] && g.type != _$_669e[303] && g.type != _$_669e[304] && 0 != g.process && (!0 === a || g.type != _$_669e[305] && g.type != _$_669e[306] && g.type != _$_669e[307]))
                    if (g.type == _$_669e[308] || g.type == _$_669e[301]) d = e[0][_$_669e[56]](g, d, a);
                    else {
                        var f;
                        (f = this.records[b][c].element) ? d = 1 == this.parent.config.stripHTML ?
                            f.innerText : f.children.length && f.children[0].classList.contains(_$_669e[27]) ? f.children[0].innerHTML : f.innerHTML: ((_$_669e[3] + d).substr(0, 1) == _$_669e[18] && (d = this.executeFormula(d, c, b, !1)), e[0] && typeof e[0][_$_669e[56]] == _$_669e[110] && (d = e[0][_$_669e[56]](g, d, a)))
                    } return d
            };
            h.setValueChained = function(c) {
                for (var b, a = [], d, e = 0; e < c.length; e++) b = c[e], d = H.getColumnNameFromCoords(b.x, b.y), a[d] = b.value;
                a = Y.getChain.call(this, a);
                var g = [],
                    f = Object.keys(a);
                for (e = 0; e < f.length; e++) 0 <= f[e].indexOf(_$_669e[39]) ?
                    (b = f[e].split(_$_669e[39]), d = b.pop(), b = I(b.join(_$_669e[39]))) : (d = f[e], b = this), d = H.getCoordsFromColumnName(d), c = U.update.call(b, d[0], d[1], a[f[e]], !0), c.w = b, c.value = c.v = b.records[d[1]][d[0]].v, c.cell = f[e], c.f = a[f[e]], g.push(c);
                g.length && G.call(this.parent, _$_669e[309], this, g)
            };
            return h
        }(),
        U = function() {
            var h = function(c) {
                c.getCell = h[_$_669e[56]];
                c.getCellFromCoords = h[_$_669e[56]];
                c.updateCell = h.update;
                c.updateCells = h.updateAll;
                c.getSelected = h.selected;
                c.isAttached = h.attached
            };
            h.attached = function(c, b) {
                return this.records[b] &&
                    this.records[b][c] && this.records[b][c].element && this.records[b][c].element.parentNode && this.records[b][c].element.parentNode.parentNode ? !0 : !1
            };
            h[_$_669e[56]] = function(c, b) {
                typeof c == _$_669e[25] && (b = H.getCoordsFromColumnName(c), c = b[0], b = b[1]);
                c = parseInt(c);
                b = parseInt(b);
                if (this.records[b] && this.records[b][c] && this.records[b][c].element) return this.records[b][c].element;
                var a = R.call(this, c, b);
                null === a && (a = _$_669e[3]);
                return U.create.call(this, c, b, a)
            };
            h.getFromCoords = function(c, b, a) {
                if (!c || null === c[0]) return null;
                var d = [],
                    e = c[0],
                    g = c[1],
                    f = c[2];
                c = c[3];
                for (a && (a = this.results ? H.invert(this.results) : null); g <= c; g++)
                    for (var k = e; k <= f; k++)
                        if (!a || a[g]) b ? d.push(H.getColumnNameFromCoords(k, g)) : d.push(this.records[g][k]);
                return d
            };
            h.selected = function(c, b) {
                return h.getFromCoords.call(this, this.getHighlighted(), c, b)
            };
            h.getSelectedColumns = function() {
                var c = [],
                    b = this.getHighlighted();
                if (b)
                    for (var a = b[0]; a <= b[2]; a++) c.push(a);
                return c
            };
            h.updateAll = function(c) {
                for (var b = this.getPrimaryKey(), a = [], d = 0; d < c.length; d++) void 0 !== c[d].id ?
                    (P.call(this, c[d]), a.push({
                        id: c[d].id,
                        y: c[d].y
                    }), !1 !== b && h.update.call(this, b, c[d].y, c[d].id, !0)) : a.push(h.update.call(this, c[d].x, c[d].y, c[d].value, c[d].force));
                L.call(this, _$_669e[135], [a])
            };
            h.update = function(c, b, a, d) {
                var e = this.records[b];
                if (e = e && e[c] ? e[c] : null) {
                    e = e.element;
                    var g = {
                        x: c,
                        y: b
                    };
                    if (d || !e || 1 != e.classList.contains(_$_669e[26])) {
                        d = G.call(this.parent, _$_669e[310], this, e, c, b, a);
                        if (!1 === d) return g;
                        void 0 !== d && (a = d);
                        d = pa[_$_669e[56]].call(this, c, b);
                        if (typeof d[0].updateCell == _$_669e[110]) {
                            var f =
                                d[0].updateCell(e, a, c, b, this, d[1]);
                            void 0 !== f && (a = f)
                        }
                        d[1] && typeof d[1].render == _$_669e[110] && d[1].render(e, a, c, b, this, d[1]);
                        d = R.call(this, c, b);
                        d !== a && (a = R.call(this, c, b, a), g.oldValue = d, g.value = a, e && h.applyOverflow.call(this, e, c, b), G.call(this.parent, _$_669e[311], this, e, c, b, g.value, g.oldValue))
                    }
                } else return !1;
                return g
            };
            h.create = function(c, b, a) {
                var d = document.createElement(_$_669e[185]);
                d.setAttribute(_$_669e[63], c);
                d.setAttribute(_$_669e[59], b);
                var e = pa[_$_669e[56]].call(this, c, b);
                if (e[0] && typeof e[0].createCell ==
                    _$_669e[110]) {
                    var g = e[0].createCell(d, a, c, b, this, e[1]);
                    void 0 !== g && (a = g)
                }
                R.call(this, c, b, a);
                this.records[b][c] || va.cell.call(this, c, b, a);
                this.records[b][c].element = d;
                !a && this.records[b][c].v && (d.innerText = this.records[b][c].v);
                e[1] && typeof e[1].render == _$_669e[110] && e[1].render(d, a, c, b, this, e[1]);
                h.applyProperties.call(this, d, c, b);
                h.applyOverflow.call(this, d, c, b);
                G.call(this.parent, _$_669e[312], this, d, c, b, a);
                return d
            };
            h.getAttributes = function(c, b, a) {
                a[c] || (a[c] = {});
                c == _$_669e[313] ? a[c][b] = this.getValue(b) :
                    c == _$_669e[314] ? this.merged.cells && (b = this.merged.cells[b]) && (a[c][b] = this.options.mergeCells[b]) : this.options[c] && this.options[c][b] && (a[c][b] = this.options[c][b])
            };
            h.setAttributes = function(c) {
                c && c.style && this.setStyle(c.style);
                c && c.comments && this.setComments(c.comments);
                c && c.meta && this.setMeta(c.meta);
                c && c.cells && this.setProperty(c.cells);
                c && c.mergeCells && this.setMerge(c.mergeCells);
                c && c.values && h.batch.call(this, c.values);
                c && c.formulas && h.batch.call(this, c.formulas, !0)
            };
            h.applyProperties = function(c,
                b, a) {
                var d = H.getColumnNameFromCoords(b, a),
                    e = this.options.columns,
                    g = this.options.style,
                    f = this.options.comments,
                    k = this.options.cells,
                    l = this.options.rows,
                    n = null,
                    p = !1;
                e[b] && (e[b].readonly || e[b].readOnly) && (p = !0);
                l && l[a] && (l[a].readonly || l[a].readOnly) && (p = !0);
                c.classList.contains(_$_669e[26]) && (p = !0);
                if (g && g[d]) {
                    var q = (q = c.getAttribute(_$_669e[127])) ? q + (_$_669e[262] + g[d]) : g[d];
                    c.setAttribute(_$_669e[127], q)
                }
                f && f[d] && (typeof f[d] === _$_669e[25] ? c.setAttribute(_$_669e[315], f[d]) : c.classList.add(_$_669e[316]));
                if (k && k[d])
                    if (!1 === k[d].locked && !0 === this.options.selectUnLockedCells ? c.classList.add(_$_669e[317]) : c.classList.remove(_$_669e[317]), k[d].zebra ? c.classList.add(_$_669e[318]) : c.classList.remove(_$_669e[318]), typeof k[d].wrap !== _$_669e[21] && ((n = k[d].wrap) ? c.classList.add(_$_669e[319]) : c.classList.remove(_$_669e[319])), !0 === k[d].readonly || !0 === k[d].readOnly) p = !0;
                    else if (!1 === k[d].readonly || !1 === k[d].readOnly) p = !1;
                null === n && (e[b] && e[b].wrap ? c.classList.add(_$_669e[319]) : c.classList.remove(_$_669e[319]));
                c.style.textAlign || (g = null, k && k[d] && k[d].align ? g = k[d].align : e && e[b].align && (g = e[b].align), g && (c.style.textAlign = g));
                void 0 == this.records[a][b].readonly && p && (this.records[a][b].readonly = p);
                this.records[a][b].readonly ? c.classList.add(_$_669e[26]) : c.classList.remove(_$_669e[26]);
                !1 === e[b].visible && c.classList.add(_$_669e[320])
            };
            h.applyOverflow = function(c, b, a) {
                1 == this.options.textOverflow && 0 < b && (b == this.options.columns.length - 1 && (c.style.overflow = _$_669e[119]), this.records[a][b - 1] && this.records[a][b -
                    1
                ].element && (this.records[a][b - 1].element.style.overflow = this.options.data[a][b] == _$_669e[3] && this.headers[b - 1] && this.headers[b - 1].offsetWidth ? _$_669e[3] : _$_669e[119]))
            };
            h.batch = function(c, b) {
                var a = [],
                    d = Object.keys(c);
                if (d.length)
                    for (var e = 0; e < d.length; e++) {
                        if (0 <= d[e].indexOf(_$_669e[39])) {
                            var g = d[e].split(_$_669e[39]);
                            var f = I(g[0]);
                            var k = g[1];
                            g = g[0]
                        } else f = this, k = d[e], g = this.getWorksheetName();
                        k = H.getCoordsFromColumnName(k);
                        k = f.updateCell(k[0], k[1], c[d[e]], !0);
                        a[g] || (a[g] = []);
                        a[g].push(k)
                    }
                if (1 ==
                    b && (d = Object.keys(a), d.length))
                    for (e = 0; e < d.length; e++) f = I(d[e]), L.call(f, _$_669e[136], a[d[e]])
            };
            return h
        }(),
        Ma = function() {
            var h = function(b) {
                b.getColumnOptions = h[_$_669e[56]];
                b.setColumnOptions = h[_$_669e[11]];
                b.getCells = h.getCells;
                b.setCells = h.setCells;
                b.getProperties = h[_$_669e[56]];
                b.setProperties = h[_$_669e[11]];
                b.getProperty = h[_$_669e[56]];
                b.setProperty = h[_$_669e[11]];
                b.resetProperty = h[_$_669e[11]];
                b.updateProperty = h.update;
                b.getOptions = h[_$_669e[56]]
            };
            h.getCells = function(b) {
                return b ? this.options.cells[b] :
                    this.options.cells
            };
            h.setCells = function(b, a) {
                if (!T.call(this.parent, this)) return !1;
                if (typeof b == _$_669e[25]) this.options.cells[b] = a;
                else {
                    a = Object.keys(b);
                    for (var d = 0; d < a.length; d++) this.options.cells[a[d]] = b[a[d]]
                }
            };
            h[_$_669e[56]] = function(b, a) {
                var d = null;
                void 0 !== a && (a = H.getColumnNameFromCoords(b, a), this.options.cells.hasOwnProperty(a) && (d = this.options.cells[a]));
                d || (d = this.options.columns[b]);
                if (!d) d = {
                    type: _$_669e[69]
                };
                else if (!d.type || typeof d.type == _$_669e[25] && !r.editors[d.type]) d.type = _$_669e[69];
                return d
            };
            var c = function(b, a, d, e, g) {
                var f = this.records[a][b].element;
                f && (e && typeof e.destroyCell == _$_669e[110] && e.destroyCell(f, b, a, this), d && typeof d.createCell == _$_669e[110] && (e = R.call(this, b, a), e = d.createCell(f, e, b, a, this, g), void 0 !== e && R.call(this, b, a, e), U.applyProperties.call(this, f, b, a)))
            };
            h[_$_669e[11]] = function(b, a, d) {
                if (!T.call(this.parent, this)) return !1;
                this.edition && this.closeEditor(this.edition, !1);
                var e = [];
                if (Array.isArray(b)) e = b;
                else if (jSuites.isNumeric(b) && jSuites.isNumeric(a)) b = {
                    x: b,
                    y: a
                }, typeof d == _$_669e[109] && (b.value = d), e.push(b);
                else if (typeof a === _$_669e[109]) b = {
                    x: b,
                    value: a
                }, e.push(b);
                else if (typeof b == _$_669e[109] && void 0 === a)
                    for (a = Object.keys(b), d = 0; d < a.length; d++) {
                        var g = H.getCoordsFromColumnName(a[d]);
                        e.push({
                            x: g[0],
                            y: g[1],
                            value: b[a[d]]
                        })
                    }
                if (Array.isArray(e))
                    for (d = 0; d < e.length; d++) {
                        var f = e[d];
                        b = f.x;
                        var k = f.y,
                            l = h[_$_669e[56]].call(this, b, k);
                        a = f.value;
                        g = r.editors[l.type];
                        var n = null;
                        a && a.type && (n = r.editors[a.type]);
                        if (typeof k === _$_669e[21])
                            for (f.oldValue = l, parseInt(a.width) ||
                                (a.width = this.options.defaultColWidth || 100), a.type && r.editors[a.type] || (a.type = _$_669e[69]), this.options.columns[b] = a, Fa.applyProperties.call(this, b), f = 0; f < this.records.length; f++) c.call(this, b, f, n, g, a);
                        else l = H.getColumnNameFromCoords(b, k), f.oldValue = this.options.cells && this.options.cells[l] ? this.options.cells[l] : _$_669e[3], a ? (this.options.cells || (this.options.cells = {}), this.options.cells[l] = a) : (delete this.options.cells[l], a = h[_$_669e[56]].call(this, b, k), n = r.editors[a.type]), c.call(this, b, k, n, g,
                            a)
                    } else return console.error(_$_669e[299]), !1;
                this.refreshBorders();
                Q.call(this.parent, {
                    worksheet: this,
                    action: _$_669e[98],
                    records: e
                });
                L.call(this, _$_669e[98], {
                    data: e
                });
                G.call(this.parent, _$_669e[321], this, e);
                return !0
            };
            h.update = function(b, a, d, e) {
                a = typeof b == _$_669e[25] || typeof b == _$_669e[22] ? [{
                    x: b,
                    y: a,
                    value: d
                }] : b;
                d = JSON.parse(JSON.stringify(a));
                if (Array.isArray(a))
                    for (var g = 0; g < a.length; g++) {
                        var f = a[g].x,
                            k = a[g].y,
                            l = a[g],
                            n = e,
                            p = H.getColumnNameFromCoords(f, k);
                        (b = this.options.cells[p]) ? l.oldValue = Object.assign({},
                            b): (l.oldValue = null, b = this.options.cells[p] = Object.assign({}, this.options.columns[f]));
                        p = Object.keys(l.value);
                        for (var q = 0; q < p.length; q++) b[p[q]] = l.value[p[q]];
                        n || typeof k === _$_669e[21] || (n = h[_$_669e[56]].call(this, f, k), n = r.editors[n.type], l.value.type ? (l = l.value, p = null, l && l.type && (p = r.editors[l.type]), c.call(this, f, k, p, n, b)) : (l = this.records[k][f].element) && typeof n.updateCell == _$_669e[110] && (p = R.call(this, f, k), p = n.updateCell(l, p, f, k, this, b), void 0 !== p && R.call(this, f, k, p), U.applyProperties.call(this,
                            l, f, k)))
                    }
                Q.call(this.parent, {
                    worksheet: this,
                    action: _$_669e[99],
                    records: a
                });
                L.call(this, _$_669e[99], {
                    data: d
                });
                G.call(this.parent, _$_669e[321], this, a)
            };
            return h
        }(),
        aa = function() {
            var h = function(c) {
                c.getRow = h[_$_669e[56]];
                c.moveRow = h.move;
                c.insertRow = h.add;
                c.deleteRow = h.del;
                c.getSelectedRows = h.selected;
                c.showRow = function(b) {
                    h.setVisibility.call(c, b, !0)
                };
                c.hideRow = function(b) {
                    h.setVisibility.call(c, b, !1)
                }
            };
            h.visibility = function(c, b) {
                Array.isArray(c) || (c = [c]);
                for (var a, d, e = b ? _$_669e[3] : _$_669e[52], g = Object.keys(c),
                        f = [], k = 0; k < g.length; k++)
                    if (a = c[g[k]], d = this.rows[a]) void 0 === d.visible && (d.visible = !0), d.visible != b && (d.element && (d.element.style.display = e, b || (d.element.remove(), -1 !== this.visibleRows.indexOf(a) && this.visibleRows.splice(a, 1))), d.visible = b, f.push(a));
                if (f.length) return J.call(this, !0), b || this.resetSelection(), this.refreshBorders(), f
            };
            h.setVisibility = function(c, b) {
                if (c = h.visibility.call(this, c, b)) Q.call(this.parent, {
                    worksheet: this,
                    action: b ? _$_669e[83] : _$_669e[82],
                    records: c
                }), L.call(this, b ? _$_669e[83] :
                    _$_669e[82], {
                        data: c
                    }), G.call(this.parent, _$_669e[322], this, b, c)
            };
            h.height = function(c) {
                var b = 0,
                    a = this.rows,
                    d = this.options.rows || {};
                a[c] && !1 !== a[c].visible && (b = a[c].height ? a[c].height : d[c] && d[c].height ? parseInt(d[c].height) : a[c].element && a[c].element.offsetHeight ? a[c].element.offsetHeight : this.options.defaultRowHeight || r.defaultRowHeight || 24);
                a[c].height || (a[c].height = b);
                return b
            };
            h[_$_669e[56]] = function(c) {
                return this.rows[c]
            };
            h.attached = function(c) {
                return this.rows[c] && this.rows[c].element && this.rows[c].element.parentNode ?
                    !0 : !1
            };
            h.selected = function(c) {
                var b = [],
                    a = this.getHighlighted();
                if (a)
                    for (var d = a[1]; d <= a[3]; d++) this.rows[d] && (c ? b.push(d) : b.push(this.rows[d].element));
                return b
            };
            h.create = function(c, b) {
                if (this.rows[c] && this.rows[c].element) return this.rows[c].element;
                this.rows[c] || va.row.call(this, c);
                var a = document.createElement(_$_669e[183]);
                a.setAttribute(_$_669e[59], c);
                this.rows[c].element = a;
                var d = null,
                    e = this.options.rows;
                e[c] && !b && (e[c].style && a.setAttribute(_$_669e[127], e[c].style), e[c].height && (a.style.height =
                    parseInt(e[c].height) + _$_669e[57]), e[c].title && (d = e[c].title), e[c].id && (this.rows[c].id = e[c].id), e[c].group && (this.rows[c].group = e[c].group, this.rows[c].state = !!e[c].state));
                !a.style.height && this.options.defaultRowHeight && (a.style.height = parseInt(this.options.defaultRowHeight) + _$_669e[57]);
                0 == this.rows[c].visible && (a.style.display = _$_669e[52]);
                d || (d = parseInt(c + 1));
                b = document.createElement(_$_669e[185]);
                b.innerHTML = d;
                b.setAttribute(_$_669e[59], c);
                b.className = _$_669e[323];
                a.appendChild(b);
                this.rows[c].group &&
                    Ja.icon.call(this, c);
                G.call(this.parent, _$_669e[324], this, c, a);
                return a
            };
            h.move = function(c, b) {
                if (!T.call(this.parent, this)) return !1;
                c = parseInt(c);
                b = parseInt(b);
                if (!this.rows[b] || !this.rows[c]) return console.error(_$_669e[325]), !1;
                if (0 < Object.keys(this.getMerge()).length) {
                    var a = null;
                    this.merged.rows[c] ? a = !0 : this.merged.rows[b] && (c > b ? this.merged.rows[b - 1] && (a = !0) : this.merged.rows[b + 1] && (a = !0));
                    a && this.destroyMerged()
                }
                this.rows[c].group && Ja.reset.call(this, c);
                this.resetBorders();
                aa.attached.call(this,
                    b) ? (a = aa.create.call(this, c), c > b ? this.tbody.insertBefore(a, this.rows[b].element) : this.tbody.insertBefore(a, this.rows[b].element.nextSibling)) : aa.attached.call(this, c) && this.tbody.removeChild(this.rows[c].element);
                this.rows.splice(b, 0, this.rows.splice(c, 1)[0]);
                this.records.splice(b, 0, this.records.splice(c, 1)[0]);
                this.options.data.splice(b, 0, this.options.data.splice(c, 1)[0]);
                L.call(this, _$_669e[80], {
                    f: c,
                    t: b
                });
                Q.call(this.parent, {
                    worksheet: this,
                    action: _$_669e[80],
                    oldValue: c,
                    newValue: b
                });
                ha.references.call(this);
                J.call(this, !0, !0);
                ha.update.call(this);
                G.call(this.parent, _$_669e[326], this, c, b)
            };
            h.add = function(c, b, a, d, e) {
                if (!T.call(this.parent, this)) return !1;
                if (!this.options.allowInsertRow) return console.error(_$_669e[327]), !1;
                var g = [];
                if (0 < c) var f = parseInt(c);
                else f = 1, Array.isArray(c) && !d && (Array.isArray(c[0]) ? d = c : (d = [], d.push(c)));
                d = Ka.standardize(d);
                a = a ? !0 : !1;
                c = this.options.data ? this.options.data.length - 1 : 0;
                if (void 0 == b || b > c || 0 > b) b = c, a = !1;
                f || (f = 1);
                if (!1 === G.call(this.parent, _$_669e[328], this, b, f, a)) return console.log(_$_669e[329]),
                    !1;
                ca.reset.call(this);
                c = a ? b : b + 1;
                for (var k, l, n = 0, p = [], q = [], v = this.rows.splice(c), w = this.options.data ? this.options.data.splice(c) : [], x = this.records.splice(c), y = c; y < f + c; y++) {
                    var B = d[n] && d[n].id ? d[n].id : za.next.call(this);
                    va.row.call(this, y, B);
                    l = this.dataType ? {} : [];
                    for (var A = 0; A < this.options.columns.length; A++) k = ra.call(this, A), l[k] = _$_669e[3], d && d[n] && d[n].data && (typeof d[n].data[k] !== _$_669e[21] ? l[k] = d[n].data[k] : typeof d[n].data[A] !== _$_669e[21] && (l[k] = d[n].data[A])), g.push({
                        x: A,
                        y,
                        value: l[k],
                        force: !0
                    });
                    q.push({
                        id: B,
                        data: l
                    });
                    n++
                }
                this.rows = this.rows.concat(v);
                this.records = this.records.concat(x);
                Array.prototype.push.apply(this.options.data, w);
                for (y = c; y < f + c; y++) this.results && (this.rows[y].results = !0), p.push(this.records[y]);
                this.merged.rows[c] && sa.updateConfig.call(this, 1, 1, c - 1, f);
                L.call(this, _$_669e[76], {
                    numOfRows: f,
                    rowNumber: b,
                    insertBefore: a ? 1 : 0,
                    data: q
                });
                this.results && (this.results = Ba.refresh.call(this));
                Q.call(this.parent, {
                    worksheet: this,
                    action: _$_669e[76],
                    rowNumber: b,
                    numOfRows: f,
                    insertBefore: a,
                    data: q
                });
                ha.references.call(this);
                J.call(this, !0, !0);
                e || ka[_$_669e[11]].call(this, 0, c, this.options.columns.length - 1, f + c - 1);
                g.length && qa.applyValues.call(this, g);
                ha.update.call(this);
                G.call(this.parent, _$_669e[330], this, b, f, p, a)
            };
            h.del = function(c, b) {
                if (!T.call(this.parent, this)) return !1;
                if (!this.options.allowDeleteRow) return console.error(_$_669e[331]), !1;
                var a = {},
                    d = [];
                void 0 == c && (b = this.getSelectedRows(!0), b.length ? (c = parseInt(b[0]), b = parseInt(b.length)) : (c = this.options.data.length - 1, b = 1));
                var e =
                    this.options.data.length - 1;
                if (void 0 == c || c > e || 0 > c) c = e;
                b || (b = 1);
                c + b >= this.options.data.length && (b = this.options.data.length - c, b >= this.options.data.length && (b = this.options.data.length, this.resetSelection()));
                if (!1 === G.call(this.parent, _$_669e[332], this, c, b)) return console.log(_$_669e[333]), !1;
                if (-1 < parseInt(c) && 0 < parseInt(b)) {
                    ca.reset.call(this);
                    for (var g = c; g < c + b; g++)
                        for (var f = 0; f < this.options.columns.length; f++) e = H.getColumnNameFromCoords(f, g), U.getAttributes.call(this, _$_669e[334], e, a), U.getAttributes.call(this,
                            _$_669e[127], e, a), U.getAttributes.call(this, _$_669e[335], e, a), U.getAttributes.call(this, _$_669e[336], e, a), U.getAttributes.call(this, _$_669e[314], e, a), U.getAttributes.call(this, _$_669e[313], e, a), d[e] = !0;
                    if (e = Y.getChain.call(this, a.values)) a.formulas = e;
                    for (g = c; g < c + b; g++) aa.attached.call(this, g) && this.rows[g].element.remove(), this.options.rows && this.options.rows[g] && (this.options.rows[g] = null);
                    e = [];
                    for (g = c; g < c + b; g++) e.push({
                        id: this.getRowId(g),
                        row: g,
                        data: this.getRowData(g)
                    });
                    this.rows.splice(c, b);
                    g =
                        this.records.splice(c, b);
                    this.options.data.splice(c, b);
                    sa.updateConfig.call(this, 1, 0, c, b);
                    this.results && (this.results = Ba.refresh.call(this));
                    L.call(this, _$_669e[77], {
                        rowNumber: c,
                        numOfRows: b,
                        data: e
                    });
                    Q.call(this.parent, {
                        worksheet: this,
                        action: _$_669e[77],
                        rowNumber: c,
                        numOfRows: b,
                        insertBefore: 1,
                        data: e,
                        attributes: a
                    });
                    ha.references.call(this, d);
                    J.call(this, !0, !0);
                    ha.update.call(this, !0);
                    G.call(this.parent, _$_669e[337], this, c, b, g, e, a)
                }
            };
            h.isVisible = function(c) {
                var b = this.rows;
                return b && !1 !== b[c].visible ?
                    !0 : !1
            };
            return h
        }(),
        na = function() {
            var h = function(c) {
                c.getColumn = h[_$_669e[56]];
                c.getColumnIdByName = h.getByName;
                c.getPrimaryKey = h.getPrimaryKey;
                c.getSelectedColumns = h.selected;
                c.moveColumn = h.move;
                c.insertColumn = h.add;
                c.deleteColumn = h.del;
                c.showColumn = function(b) {
                    h.setVisibility.call(c, b, !0)
                };
                c.hideColumn = function(b) {
                    h.setVisibility.call(c, b, !1)
                }
            };
            h.visibility = function(c, b) {
                Array.isArray(c) || (c = [c]);
                for (var a, d, e, g = Object.keys(c), f = [], k = 0; k < g.length; k++)
                    if (a = c[g[k]], (d = this.options.columns[a]) && d.type !==
                        _$_669e[119] && (void 0 === d.visible && (d.visible = !0), d.visible != b)) {
                        if (this.colgroup[a])
                            for (e = b ? parseInt(d.width || this.options.defaultColWidth) : _$_669e[126], this.colgroup[a].setAttribute(_$_669e[196], e), e = 0; e < this.rows.length; e++) this.records[e] && this.records[e][a] && this.records[e][a].element && (b ? this.records[e][a].element.classList.remove(_$_669e[320]) : this.records[e][a].element.classList.add(_$_669e[320]));
                        d.visible = b;
                        f.push(a)
                    } if (f.length) return J.call(this, !0), b || this.resetSelection(), this.refreshBorders(),
                    f
            };
            h.setVisibility = function(c, b) {
                if (c = h.visibility.call(this, c, b)) Q.call(this.parent, {
                    worksheet: this,
                    action: b ? _$_669e[85] : _$_669e[84],
                    records: c
                }), L.call(this, b ? _$_669e[85] : _$_669e[84], {
                    data: c
                }), G.call(this.parent, _$_669e[338], this, b, c)
            };
            h[_$_669e[56]] = function(c) {
                return this.options.columns[c] || !1
            };
            h.attached = function(c) {
                return this.headers[c] && this.headers[c].parentNode ? !0 : !1
            };
            h.getByName = function(c) {
                for (var b = this.options.columns, a = 0; a < b.length; a++)
                    if (b[a].name == c) return a;
                return !1
            };
            h.getPrimaryKey =
                function() {
                    for (var c = 0; c < this.options.columns.length; c++)
                        if (1 == this.options.columns[c].primaryKey) return c;
                    return !1
                };
            h.selected = function() {
                var c = [],
                    b = this.getHighlighted();
                if (b)
                    for (var a = b[0]; a <= b[2]; a++) c.push(a);
                return c
            };
            h.move = function(c, b) {
                if (!T.call(this.parent, this)) return !1;
                c = parseInt(c);
                b = parseInt(b);
                if (!this.options.columns[c]) return console.error(_$_669e[339]), !1;
                if (!this.options.columns[b]) return console.error(_$_669e[340]), !1;
                if (0 < Object.keys(this.getMerge()).length) {
                    var a = null;
                    this.merged.cols[c] ?
                        a = !0 : this.merged.cols[b] && (c > b ? this.merged.cols[b - 1] && (a = !0) : this.merged.cols[b + 1] && (a = !0));
                    a && this.destroyMerged()
                }
                ca.reset.call(this);
                this.options.freezeColumns && ea.applyReset.call(this);
                if (na.attached.call(this, b)) {
                    Fa.create.call(this, c);
                    var d = (c > b ? this.headers[b] : this.headers[b].nextSibling) || null;
                    this.headerContainer.insertBefore(this.headers[c], d);
                    d = (c > b ? this.colgroup[b] : this.colgroup[b].nextSibling) || null;
                    this.colgroupContainer.insertBefore(this.colgroup[c], d);
                    for (var e = 0; e < this.tbody.children.length; e++) a =
                        this.tbody.children[e].getAttribute(_$_669e[59]), d = (c > b ? this.records[a][b].element : this.records[a][b].element.nextSibling) || null, this.tbody.children[e].insertBefore(U[_$_669e[56]].call(this, c, a), d);
                    if (this.options.footers)
                        for (e = 0; e < this.tfoot.children.length; e++) d = (c > b ? this.footers.content[e][b].element : this.footers.content[e][b].element.nextSibling) || null, this.tfoot.children[e].insertBefore(ia.create.call(this, c, e), d)
                } else if (na.attached.call(this, c)) {
                    this.headerContainer.removeChild(this.headers[c]);
                    this.colgroupContainer.removeChild(this.colgroup[c]);
                    for (e = 0; e < this.tbody.children.length; e++) a = parseInt(this.tbody.children[e].getAttribute(_$_669e[59])), this.tbody.children[e].removeChild(this.records[a][c].element);
                    if (this.options.footers)
                        for (e = 0; e < this.tfoot.children.length; e++) this.tfoot.children[e].removeChild(this.footers.content[e][c].element)
                }
                this.options.columns.splice(b, 0, this.options.columns.splice(c, 1)[0]);
                this.headers.splice(b, 0, this.headers.splice(c, 1)[0]);
                this.colgroup.splice(b, 0,
                    this.colgroup.splice(c, 1)[0]);
                for (e = 0; e < this.rows.length; e++) this.dataType || this.options.data[e].splice(b, 0, this.options.data[e].splice(c, 1)[0]), this.records[e].splice(b, 0, this.records[e].splice(c, 1)[0]);
                if (this.options.footers)
                    for (e = 0; e < this.options.footers.length; e++) this.options.footers[e].splice(b, 0, this.options.footers[e].splice(c, 1)[0]), this.footers.content[e].splice(b, 0, this.footers.content[e].splice(c, 1)[0]);
                L.call(this, _$_669e[81], {
                    f: c,
                    t: b
                });
                Q.call(this.parent, {
                    worksheet: this,
                    action: _$_669e[81],
                    oldValue: c,
                    newValue: b
                });
                ha.references.call(this);
                J.call(this, !0, !0);
                ha.update.call(this);
                G.call(this.parent, _$_669e[341], this, c, b)
            };
            h.add = function(c, b, a, d, e, g, f) {
                if (!T.call(this.parent, this)) return !1;
                if (!this.options.allowInsertColumn) return console.error(_$_669e[342]), !1;
                var k = [];
                if (0 < c) var l = parseInt(c);
                else if (l = 1, Array.isArray(c) && !e)
                    if (Array.isArray(c[0])) e = c;
                    else {
                        e = [];
                        for (var n = 0; n < c.length; n++) e[n] = [c[n]]
                    } e || (e = []);
                a = a ? !0 : !1;
                c = this.options.columns.length - 1;
                if (void 0 == b || b > c || 0 > b) b = c, a = !1;
                l ||
                    (l = 1);
                if (!1 === G.call(this.parent, _$_669e[343], this, b, l, a)) return console.log(_$_669e[344]), !1;
                ca.reset.call(this);
                this.options.freezeColumns && ea.applyReset.call(this);
                if (1 == this.dataType) var p = parseInt((_$_669e[3] + Date.now()).substr(-8));
                d || (d = []);
                for (n = 0; n < l; n++) d[n] || (d[n] = {}), d[n].type || (d[n].type = _$_669e[69]), d[n].width || (d[n].width = this.options.defaultColWidth), d[n].type != _$_669e[71] && d[n].type != _$_669e[345] || d[n].source || (d[n].source = []), 1 == this.dataType && typeof d[n].name == _$_669e[21] && (d[n].name =
                    _$_669e[346] + (p + n));
                p = a ? b : b + 1;
                this.options.columns = H.injectArray(this.options.columns, p, d);
                n = this.headers.splice(p);
                var q = this.colgroup.splice(p);
                for (c = p; c < l + p; c++) Fa.create.call(this, c);
                this.headers = this.headers.concat(n);
                this.colgroup = this.colgroup.concat(q);
                ia.adjust.call(this, p, l, 1);
                Aa.adjust.call(this, b, l, 1, g);
                q = null;
                for (var v, w = 0; w < this.options.data.length; w++) {
                    this.dataType || (q = this.options.data[w].splice(p));
                    v = this.records[w].splice(p);
                    g = 0;
                    for (c = p; c < l + p; c++) n = e[w] && e[w][g] ? e[w][g] : _$_669e[3],
                        R.call(this, c, w, n), k.push({
                            x: c,
                            y: w,
                            value: n
                        }), va.cell.call(this, c, w), g++;
                    this.dataType || (this.options.data[w] = this.options.data[w].concat(q));
                    this.records[w] = this.records[w].concat(v)
                }
                this.merged.cols[p] && sa.updateConfig.call(this, 0, 1, p - 1, l);
                L.call(this, _$_669e[78], {
                    numOfColumns: l,
                    columnNumber: b,
                    insertBefore: a ? 1 : 0,
                    properties: d,
                    data: e
                });
                Q.call(this.parent, {
                    worksheet: this,
                    action: _$_669e[78],
                    columnNumber: b,
                    numOfColumns: l,
                    insertBefore: a,
                    properties: d,
                    data: e
                });
                ha.references.call(this);
                J.call(this, !0, !0);
                f || ka[_$_669e[11]].call(this, p, 0, l + p - 1, this.rows.length - 1);
                k.length && qa.applyValues.call(this, k);
                ha.update.call(this);
                G.call(this.parent, _$_669e[347], this, b, l, [], a)
            };
            h.del = function(c, b) {
                if (!T.call(this.parent, this)) return !1;
                if (!this.options.allowDeleteColumn) return console.error(_$_669e[348]), !1;
                var a = {},
                    d = [];
                if (1 < this.options.columns.length) {
                    void 0 == c && (b = this.getSelectedColumns(), b.length ? (c = parseInt(b[0]), b = parseInt(b.length)) : (c = this.options.columns.length - 1, b = 1));
                    var e = this.options.columns.length -
                        1;
                    if (void 0 == c || c > e || 0 > c) c = e;
                    b || (b = 1);
                    c + b >= this.options.columns.length && (b = this.options.columns.length - c, b >= this.options.columns.length && (b--, console.error(_$_669e[349])));
                    if (!1 === G.call(this.parent, _$_669e[350], this, c, b)) return console.log(_$_669e[351]), !1;
                    if (-1 < parseInt(c) && 0 < parseInt(b)) {
                        ca.reset.call(this);
                        e = null;
                        var g = [],
                            f = [],
                            k = [];
                        var l = 0;
                        for (var n = c; n < c + b; n++) na.attached.call(this, n) && (this.colgroup[n].parentNode.removeChild(this.colgroup[n]), this.headers[n].parentNode.removeChild(this.headers[n])),
                            k[l] = this.options.columns[n], l++;
                        for (var p = 0; p < this.options.data.length; p++)
                            for (n = c; n < c + b; n++) l = H.getColumnNameFromCoords(n, p), U.getAttributes.call(this, _$_669e[334], l, a), U.getAttributes.call(this, _$_669e[127], l, a), U.getAttributes.call(this, _$_669e[335], l, a), U.getAttributes.call(this, _$_669e[336], l, a), U.getAttributes.call(this, _$_669e[314], l, a), U.getAttributes.call(this, _$_669e[313], l, a), d[l] = !0, this.records[p][n].element && 0 <= Array.prototype.indexOf.call(this.rows[p].element.children, this.records[p][n].element) &&
                                this.records[p][n].element.parentNode.removeChild(this.records[p][n].element);
                        if (n = Y.getChain.call(this, a.values)) a.formulas = n;
                        this.options.columns.splice(c, b);
                        this.headers.splice(c, b);
                        this.colgroup.splice(c, b);
                        for (p = 0; p < this.options.data.length; p++) this.dataType || (f[p] = this.options.data[p].splice(c, b)), g[p] = this.records[p].splice(c, b);
                        ia.adjust.call(this, c, b, 0);
                        (n = Aa.adjust.call(this, c, b, 0)) && (e = {
                            nested: n
                        });
                        sa.updateConfig.call(this, 0, 0, c, b);
                        L.call(this, _$_669e[79], {
                            columnNumber: c,
                            numOfColumns: b
                        });
                        Q.call(this.parent, {
                            worksheet: this,
                            action: _$_669e[79],
                            columnNumber: c,
                            numOfColumns: b,
                            insertBefore: 1,
                            data: f,
                            properties: k,
                            attributes: a,
                            extra: e
                        });
                        ha.references.call(this, d);
                        J.call(this, !0, !0);
                        ha.update.call(this, !0);
                        G.call(this.parent, _$_669e[352], this, c, b, g, k, a)
                    }
                }
            };
            h.isVisible = function(c) {
                var b = this.options.columns;
                return !1 === b[c].visible || b[c].type === _$_669e[119] || b[c].width === _$_669e[173] || b[c].width === _$_669e[126] ? !1 : !0
            };
            return h
        }(),
        Na = function() {
            var h = {},
                c = {},
                b = function(d) {
                    if (typeof d == _$_669e[25]) {
                        var e =
                            d;
                        d = H.getCoordsFromColumnName(d);
                        var g = d[0];
                        d = d[1]
                    } else g = d.x, d = d.y, e = H.getColumnNameFromCoords(g, d);
                    if (this.records[d] && this.records[d][g] && this.records[d][g].element) {
                        c[e] = this.options.style[e];
                        this.records[d][g].element.setAttribute(_$_669e[127], _$_669e[3]);
                        if (this.options.cells && this.options.cells[e] && this.options.cells[e].align) var f = this.options.cells[e].align;
                        else this.options.columns[g].align && (f = this.options.columns[g].align);
                        f && (this.records[d][g].element.style.textAlign = f)
                    }
                    this.options.style &&
                        this.options.style[e] && delete this.options.style[e]
                },
                a = function(d) {
                    d.getStyle = a[_$_669e[56]];
                    d.setStyle = a[_$_669e[11]];
                    d.resetStyle = a.reset
                };
            a[_$_669e[56]] = function(d) {
                var e = this.options.style;
                if (!e) return !1;
                if (d) {
                    if (typeof d == _$_669e[25]) return e[d];
                    var g = [],
                        f = Object.keys(d);
                    if (0 < f.length)
                        for (var k = 0; k < f.length; k++) {
                            if (typeof d[k] == _$_669e[25]) var l = e && e[d[k]] ? e[d[k]] : _$_669e[3];
                            else l = H.getColumnNameFromCoords(d[k].x, d[k].y), l = e && e[l] ? e[l] : _$_669e[3];
                            g.push(l)
                        }
                    return g
                }
                return this.options.style
            };
            a.getFromSelection = function(d) {
                var e, g, f = this.options.style,
                    k = 0,
                    l = [];
                for (g = d[1]; g <= d[3]; g++) {
                    var n = 0;
                    l[k] = [];
                    for (e = d[0]; e <= d[2]; e++) {
                        var p = H.getColumnNameFromCoords(e, g);
                        l[k][n++] = f && f[p] ? f[p] : _$_669e[3]
                    }
                    k++
                }
                return l
            };
            a[_$_669e[11]] = function(d, e, g, f) {
                if (!T.call(this.parent, this)) return !1;
                this.options.style || (this.options.style = {});
                var k = this.options.style,
                    l = null,
                    n = [],
                    p = [],
                    q = {};
                p = function(x, y, B) {
                    if (!q[x] && (k[x] ? n = k[x].split(_$_669e[262]) : (k[x] = _$_669e[3], n = []), q[x] = {}, n.length))
                        for (var A = 0; A < n.length; A++) n[A] &&
                            n[A].trim() && (l = n[A].split(_$_669e[164]), q[x][l[0].trim()] = l[1].trim());
                    c[x] || (c[x] = []);
                    h[x] || (h[x] = []);
                    q[x][y] || (q[x][y] = _$_669e[3]);
                    c[x].push([y + _$_669e[164] + q[x][y]]);
                    q[x][y] = q[x][y] && q[x][y] == B && !f ? _$_669e[3] : B;
                    h[x].push([y + _$_669e[164] + q[x][y]]);
                    B = H.getCoordsFromColumnName(x);
                    this.records[B[1]] && this.records[B[1]][B[0]] && this.records[B[1]][B[0]].element && (this.records[B[1]][B[0]].element.style[y] = q[x][y]);
                    q[x][y] || delete q[x][y]
                };
                h = {};
                c = {};
                if (e)
                    if (typeof d == _$_669e[25]) p.call(this, d, e, g);
                    else {
                        if (d &&
                            d.length)
                            for (var v = 0; v < d.length; v++) {
                                var w = H.getColumnNameFromCoords(d[v].x, d[v].y);
                                p.call(this, w, e, g)
                            }
                    }
                else
                    for (g = Object.keys(d), v = 0; v < g.length; v++)
                        if (w = d[g[v]])
                            for (typeof w == _$_669e[25] && (w = w.split(_$_669e[262])), e = 0; e < w.length; e++) w[e] && (typeof w[e] == _$_669e[25] && (w[e] = w[e].split(_$_669e[164])), w[e][0].trim() && p.call(this, g[v], w[e][0].trim(), w[e][1].trim()));
                n = {};
                p = Object.keys(q);
                if (p.length) {
                    for (v = 0; v < p.length; v++) {
                        n[p[v]] || (n[p[v]] = []);
                        d = Object.keys(q[p[v]]);
                        for (e = 0; e < d.length; e++) n[p[v]].push(d[e] +
                            _$_669e[144] + q[p[v]][d[e]]);
                        this.options.style[p[v]] = n[p[v]].join(_$_669e[353])
                    }
                    p = Object.keys(h);
                    for (v = 0; v < p.length; v++) h[p[v]] = h[p[v]].join(_$_669e[262]);
                    p = Object.keys(c);
                    for (v = 0; v < p.length; v++) c[p[v]] = c[p[v]].join(_$_669e[262]);
                    this.refreshBorders();
                    Q.call(this.parent, {
                        worksheet: this,
                        action: _$_669e[88],
                        oldValue: c,
                        newValue: h
                    });
                    L.call(this, _$_669e[88], [h]);
                    G.call(this.parent, _$_669e[354], this, h, c)
                }
            };
            a.reset = function(d) {
                if (!T.call(this.parent, this)) return !1;
                if (d) {
                    c = {};
                    if (Array.isArray(d))
                        for (var e =
                                0; e < d.length; e++) b.call(this, d[e]);
                    else b.call(this, d);
                    this.refreshBorders();
                    Q.call(this.parent, {
                        worksheet: this,
                        action: _$_669e[89],
                        cells: d,
                        oldValue: c
                    });
                    L.call(this, _$_669e[89], [d]);
                    G.call(this.parent, _$_669e[356], this, d)
                } else console.error(_$_669e[355])
            };
            a.update = function(d) {
                h = {};
                c = {};
                for (var e = Object.keys(d), g = 0; g < e.length; g++) {
                    c[e[g]] = this.options.style[e[g]] || _$_669e[3];
                    h[e[g]] = this.options.style[e[g]] = d[e[g]];
                    var f = H.getCoordsFromColumnName(e[g]),
                        k = f[0];
                    f = f[1];
                    if (this.records[f] && this.records[f][k] &&
                        this.records[f][k].element && (this.records[f][k].element.setAttribute(_$_669e[127], d[e[g]]), !this.records[f][k].element.style.textAlign)) {
                        var l = null;
                        this.options.cells && this.options.cells[e[g]] && this.options.cells[e[g]].align ? l = this.options.cells[e[g]].align : this.options.columns[k].align && (l = this.options.columns[k].align);
                        l && (this.records[f][k].element.style.textAlign = l)
                    }
                }
                L.call(this, _$_669e[89], [e]);
                L.call(this, _$_669e[88], [h]);
                G.call(this.parent, _$_669e[354], this, h, c)
            };
            return a
        }(),
        bb = function() {
            var h =
                function(c) {
                    c.getComments = h[_$_669e[56]];
                    c.setComments = h[_$_669e[11]]
                };
            h[_$_669e[56]] = function(c) {
                return c && typeof c == _$_669e[25] ? this.options.comments && this.options.comments[c] ? this.options.comments[c] : !1 : this.options.comments
            };
            h[_$_669e[11]] = function(c, b) {
                if (!T.call(this.parent, this)) return !1;
                if (typeof c == _$_669e[25]) {
                    var a = {};
                    a[c] = b
                } else a = c;
                if (b = G.call(this.parent, _$_669e[357], this, a)) a = b;
                else if (!1 === b) return !1;
                this.options.comments || (this.options.comments = {});
                c = {};
                var d = Object.keys(a);
                if (d.length) {
                    for (var e =
                            0; e < d.length; e++) {
                        var g = H.getCoordsFromColumnName(d[e]);
                        this.records[g[1]] && this.records[g[1]][g[0]] && (c[d[e]] = this.options.comments[d[e]] || _$_669e[3], (b = a[d[e]]) ? this.options.comments[d[e]] = b : delete this.options.comments[d[e]], this.records[g[1]][g[0]].element && (b ? this.records[g[1]][g[0]].element.setAttribute(_$_669e[315], b) : this.records[g[1]][g[0]].element.removeAttribute(_$_669e[315])))
                    }
                    Q.call(this.parent, {
                        worksheet: this,
                        action: _$_669e[93],
                        newValue: a,
                        oldValue: c
                    });
                    L.call(this, _$_669e[93], [a]);
                    G.call(this.parent,
                        _$_669e[358], this, a, c)
                }
            };
            return h
        }(),
        cb = function() {
            var h = function(c) {
                c.getMeta = h[_$_669e[56]];
                c.setMeta = h[_$_669e[11]];
                c.resetMeta = h.reset
            };
            h[_$_669e[56]] = function(c, b) {
                if (c) {
                    if (this.options.meta[c]) return b ? this.options.meta[c][b] : this.options.meta[c]
                } else return this.options.meta
            };
            h[_$_669e[11]] = function(c, b, a) {
                if (!T.call(this.parent, this)) return !1;
                this.options.meta || (this.options.meta = {});
                if (typeof c == _$_669e[25] && b) {
                    a || (a = _$_669e[3]);
                    var d = {};
                    d[c] = {};
                    d[c][b] = a;
                    c = d
                }
                b = Object.keys(c);
                if (b.length) {
                    a = {};
                    for (d = 0; d < b.length; d++) {
                        this.options.meta[b[d]] || (this.options.meta[b[d]] = {});
                        a[b[d]] || (a[b[d]] = {});
                        for (var e = Object.keys(c[b[d]]), g = 0; g < e.length; g++) this.options.meta[b[d]][e[g]] = c[b[d]][e[g]], a[b[d]][e[g]] = c[b[d]][e[g]]
                    }
                    Object.keys(a).length && (L.call(this, _$_669e[359], [a]), G.call(this.parent, _$_669e[360], this, a))
                }
            };
            h.reset = function(c) {
                if (!T.call(this.parent, this)) return !1;
                if (c) {
                    if (this.options.meta) {
                        typeof c == _$_669e[25] && (c = [c]);
                        for (var b = 0; b < c.length; b++) this.options.meta[c[b]] && delete this.options.meta[c[b]]
                    }
                } else c = {}, this.options.meta = c;
                L.call(this, _$_669e[361], c);
                G.call(this.parent, _$_669e[362], this, c)
            };
            return h
        }(),
        db = function() {
            var h = function(c) {
                c.orderBy = h.execute
            };
            h.handler = function(c, b) {
                return function(a, d) {
                    a = a[1];
                    d = d[1];
                    return c ? a === _$_669e[3] && d !== _$_669e[3] ? 1 : a !== _$_669e[3] && d === _$_669e[3] ? -1 : a > d ? -1 : a < d ? 1 : 0 : a === _$_669e[3] && d !== _$_669e[3] ? 1 : a !== _$_669e[3] && d === _$_669e[3] ? -1 : a > d ? 1 : a < d ? -1 : 0
                }
            };
            h.execute = function(c, b, a, d) {
                if (0 <= c) {
                    b = null == b ? this.headers[c] && this.headers[c].classList.contains(_$_669e[363]) ?
                        1 : 0 : b ? 1 : 0;
                    for (var e, g, f = [], k = 0; k < this.options.data.length; k++) e = Ma[_$_669e[56]].call(this, c, k), g = this.records[k][c].v, g = e.type == _$_669e[302] || e.type == _$_669e[306] || e.type == _$_669e[307] || e.type == _$_669e[364] || e.type == _$_669e[22] || e.type == _$_669e[365] ? Number(g) : jSuites.isNumeric(g) ? Number(g) : g.toLowerCase(), f[k] = [k, g];
                    k = typeof this.parent.config.sorting === _$_669e[110] ? this.parent.config.sorting(b, c) : h.handler(b, c);
                    f.sort(k);
                    e = a ? !0 : !1;
                    if (!a)
                        for (a = [], k = 0; k < f.length; k++) a[k] = f[k][0];
                    if (k = G.call(this.parent,
                            _$_669e[366], this, c, b, a)) a = k;
                    else if (!1 === k) return !1;
                    if (e) this.destroyMerged();
                    else if (0 < Object.keys(this.getMerge()).length)
                        if (confirm(D(_$_669e[367]))) this.destroyMerged();
                        else return !1;
                    if (a.length < f.length)
                        for (k = 0; k < f.length; k++) - 1 == a.indexOf(f[k][0]) && a.push(f[k][0]);
                    f = [];
                    for (k = 0; k < a.length; k++) f[a[k]] = k;
                    Ia.call(this, !1);
                    ca.reset.call(this);
                    L.call(this, _$_669e[96], {
                        column: c,
                        direction: b,
                        order: a
                    });
                    Q.call(this.parent, {
                        worksheet: this,
                        action: _$_669e[96],
                        column: c,
                        newValue: a,
                        oldValue: f,
                        direction: b,
                        previousState: h.getState.call(this)
                    });
                    h.update.call(this, a);
                    h.arrow.call(this, c, b, d);
                    ca.refresh.call(this);
                    G.call(this.parent, _$_669e[368], this, c, b, a);
                    return !0
                }
            };
            h.update = function(c) {
                for (var b = [], a = 0; a < c.length; a++) b[a] = this.options.data[c[a]];
                for (a = 0; a < c.length; a++) this.options.data[a] = b[a];
                b = [];
                for (a = 0; a < c.length; a++) b[a] = this.records[c[a]];
                this.records = b;
                b = [];
                for (a = 0; a < c.length; a++) b[a] = this.rows[c[a]];
                this.rows = b;
                ha.references.call(this);
                this.tbody.innerHTML = _$_669e[3];
                J.call(this, !0, !0);
                this.searchInput && this.searchInput.value ? xa.update.call(this,
                    null) : la.update.call(this)
            };
            h.arrow = function(c, b, a) {
                for (var d = 0; d < this.headers.length; d++) this.headers[d] && (this.headers[d].classList.remove(_$_669e[369]), this.headers[d].classList.remove(_$_669e[363]));
                void 0 === a ? this.headers[c] && this.headers[c].classList.add(b ? _$_669e[369] : _$_669e[363]) : Array.isArray(a) && this.headers[a[0]] && this.headers[a[0]].classList.add(a[1] ? _$_669e[369] : _$_669e[363])
            };
            h.getState = function() {
                for (var c = 0; c < this.headers.length; c++)
                    if (this.headers[c]) {
                        if (this.headers[c].classList.contains(_$_669e[369])) return [c,
                            1
                        ];
                        if (this.headers[c].classList.contains(_$_669e[363])) return [c, 0]
                    } return null
            };
            return h
        }(),
        Ea = function() {
            var h = function(c) {
                c.whichPage = h.whichPage;
                c.quantityOfPages = h.quantityOfPages;
                c.page = h[_$_669e[11]];
                c.updatePagination = h.update
            };
            h.build = function() {
                this.pageNumber = 0;
                var c = document.createElement(_$_669e[32]),
                    b = document.createElement(_$_669e[32]);
                this.pagination = document.createElement(_$_669e[32]);
                this.pagination.classList.add(_$_669e[370]);
                this.pagination.appendChild(c);
                this.pagination.appendChild(b);
                this.options.pagination || (this.pagination.style.display = _$_669e[52]);
                this.element.appendChild(this.pagination)
            };
            h.pageUp = function() {
                0 < this.pageNumber && (this.pageNumber--, h[_$_669e[11]].call(this, this.pageNumber))
            };
            h.pageDown = function() {
                this.pageNumber < h.quantityOfPages.call(this) - 1 && (this.pageNumber++, h[_$_669e[11]].call(this, this.pageNumber))
            };
            h.whichPage = function(c) {
                if (0 < this.options.pagination) return this.rows[c] ? (this.results && (c = this.results.indexOf(parseInt(c))), Math.ceil((parseInt(c) + 1) / parseInt(this.options.pagination)) -
                    1) : null;
                console.log(_$_669e[371]);
                return !1
            };
            h.quantityOfPages = function() {
                if (0 < this.options.pagination) return Math.ceil((this.results ? this.results.length : this.rows.length) / parseInt(this.options.pagination));
                console.log(_$_669e[371]);
                return !1
            };
            h[_$_669e[11]] = function(c, b) {
                var a = this.pageNumber,
                    d = parseInt(this.options.pagination);
                if (d) {
                    if (null == c || -1 == c) c = Math.ceil((this.results ? this.results : this.rows).length / d) - 1;
                    if (!1 === G.call(this.parent, _$_669e[373], this, c, a, d)) return !1;
                    this.pageNumber = c;
                    J.update.call(this);
                    ca.refresh.call(this);
                    ia.refresh.call(this);
                    typeof b == _$_669e[110] && b.call(this);
                    G.call(this.parent, _$_669e[374], this, c, a, d)
                } else console.error(_$_669e[372])
            };
            h.update = function() {
                this.pagination.children[0].innerHTML = _$_669e[3];
                this.pagination.children[1].innerHTML = _$_669e[3];
                if (this.options.pagination) {
                    var c = this.results ? this.results.length : this.rows.length;
                    if (c) {
                        c = Math.ceil(c / this.options.pagination);
                        if (6 > this.pageNumber) var b = 1,
                            a = 10 > c ? c : 10;
                        else 5 > c - this.pageNumber ? (b = c - 9, a = c, 1 > b && (b = 1)) : (b = this.pageNumber -
                            4, a = this.pageNumber + 5);
                        if (1 < b) {
                            var d = document.createElement(_$_669e[32]);
                            d.className = _$_669e[376];
                            d.innerHTML = _$_669e[15];
                            d.title = 1;
                            this.pagination.children[1].appendChild(d)
                        }
                        for (; b <= a; b++) d = document.createElement(_$_669e[32]), d.className = _$_669e[376], d.innerHTML = b, this.pagination.children[1].appendChild(d), this.pageNumber == b - 1 && d.classList.add(_$_669e[377]);
                        a < c && (d = document.createElement(_$_669e[32]), d.className = _$_669e[376], d.innerHTML = _$_669e[17], d.title = c, this.pagination.children[1].appendChild(d));
                        this.pagination.children[0].innerHTML = D(_$_669e[378], [this.pageNumber + 1, c])
                    } else this.pagination.children[0].innerHTML = D(_$_669e[375]);
                    this.options.tableOverflow && (this.content.style.height = this.table.offsetHeight + 16 + _$_669e[57], this.scroll.style.height = _$_669e[173], this.table.style.top = 0)
                }
            };
            return h
        }(),
        ja = function() {
            var h = function(c) {
                c.setFreezeRows = h[_$_669e[11]];
                c.resetFreezeRows = h.reset
            };
            h.update = function(c) {
                var b, a;
                if (a = this.options.freezeRows)
                    for (var d = this.thead.offsetHeight, e = 0; e < a; e++) {
                        if (b =
                            this.records[e][c].element) b.classList.add(_$_669e[379]), b.style.top = d + _$_669e[57];
                        d += parseInt(this.rows[e].element.offsetHeight)
                    }
            };
            h.headers = function() {
                var c = this.thead.offsetHeight,
                    b;
                if (b = this.options.freezeRows)
                    for (var a = 0; a < b; a++) this.rows[a] && this.rows[a].element && (this.rows[a].element.firstChild.classList.add(_$_669e[379]), this.rows[a].element.firstChild.style.top = c + _$_669e[57], c += parseInt(this.rows[a].element.offsetHeight));
                else h.reset.call(this)
            };
            h.refresh = function() {
                var c = this.options.freezeRows;
                if (0 < c)
                    if (h.getHeight.call(this) > this.content.offsetHeight) console.error(_$_669e[380]), this.options.freezeRows--;
                    else if (this.merged.rows[c]) console.error(_$_669e[381]), this.options.freezeRows = 0;
                else
                    for (h.headers.call(this), c = 0; c < this.visibleCols.length; c++) h.update.call(this, this.visibleCols[c])
            };
            h[_$_669e[11]] = function(c) {
                h.applyReset.call(this);
                var b = this.options.freezeRows;
                this.options.freezeRows = c;
                this.options.freezeRowControl && !c ? this.thead.lastChild.firstChild.classList.add(_$_669e[382]) :
                    this.thead.lastChild.firstChild.classList.remove(_$_669e[382]);
                h.refresh.call(this);
                J.call(this, !0);
                ca.refresh.call(this);
                Q.call(this.parent, {
                    worksheet: this,
                    action: _$_669e[95],
                    newValue: c,
                    oldValue: b
                });
                L.call(this, _$_669e[95], {
                    row: c
                })
            };
            h.reset = function() {
                h[_$_669e[11]].call(this, 0)
            };
            h.getHeight = function(c) {
                c = c ? 0 : this.thead.offsetHeight;
                if (0 < this.options.freezeRows)
                    for (var b = 0; b < this.options.freezeRows; b++) c += parseInt(aa.height.call(this, b));
                return c
            };
            h.applyReset = function() {
                for (var c, b = 0; b < this.options.freezeRows; b++)
                    if (this.rows[b].element) {
                        this.rows[b].element.firstChild.classList.remove(_$_669e[379]);
                        this.rows[b].element.firstChild.style.top = _$_669e[3];
                        for (var a = 0; a < this.options.columns.length; a++)
                            if (c = this.records[b][a].element) c.classList.remove(_$_669e[379]), c.style.top = _$_669e[3]
                    }
            };
            h.start = function(c, b) {
                c = this.parent.helper;
                h.event = {
                    y: b,
                    d: b
                };
                this.content.scrollTop = 0;
                b = this.content.offsetTop + h.getHeight.call(this);
                c.style.display = _$_669e[24];
                c.style.top = b + _$_669e[57];
                c.style.left = _$_669e[173];
                c.style.width = this.content.offsetWidth + _$_669e[57];
                c.style.height = _$_669e[200]
            };
            h.end = function(c) {
                c =
                    h.event;
                null !== c.d && this.setFreezeRows(c.d);
                this.parent.helper.style.display = _$_669e[3];
                h.event = null
            };
            h.cancel = function() {
                h.event.d = null;
                h.end.call(this, !1)
            };
            h.move = function(c) {
                var b = c.target.getAttribute(_$_669e[59]);
                if (null != b) {
                    b = parseInt(b);
                    var a = (c = c.target.clientHeight / 2 > c.offsetY) ? b : b + 1;
                    this.merged.rows[a] && 0 < a || (h.event.d = a, a = this.parent.element.getBoundingClientRect(), b = this.rows[b].element.getBoundingClientRect(), this.parent.helper.style.top = (c ? b.top : b.bottom) - a.top + _$_669e[57])
                }
            };
            return h
        }(),
        ea = function() {
            var h = function(c) {
                c.setFreezeColumns = h[_$_669e[11]];
                c.resetFreezeColumns = h.reset
            };
            h.width = function() {
                if (this.options.freezeColumns) {
                    h.headers.call(this);
                    for (var c = 0; c < this.rows.length; c++) this.rows[c].element && h.update.call(this, c)
                }
            };
            h.update = function(c) {
                var b, a;
                if (a = this.options.freezeColumns) {
                    var d = Ca.call(this) ? 50 : 0;
                    if (this.rows[c].element)
                        for (var e = 0; e < a; e++) {
                            if (b = this.records[c][e].element) b.classList.add(_$_669e[58]), b.style.left = d + _$_669e[57];
                            na.isVisible.call(this, e) && (d += parseInt(this.options.columns[e].width))
                        }
                }
            };
            h.headers = function() {
                var c = Ca.call(this) ? 50 : 0,
                    b;
                if (b = this.options.freezeColumns)
                    for (var a = 0; a < b; a++) {
                        if (this.headers[a]) {
                            this.headers[a].classList.add(_$_669e[58]);
                            this.headers[a].style.left = c + _$_669e[57];
                            if (this.options.footers)
                                for (var d = 0; d < this.options.footers.length; d++) this.footers.content[d][a].element.classList.add(_$_669e[58]), this.footers.content[d][a].element.style.left = c + _$_669e[57];
                            !1 !== this.options.columns[a].visible && (c += parseInt(this.options.columns[a].width))
                        }
                    } else h.reset.call(this)
            };
            h.refresh = function() {
                var c = this.options.freezeColumns;
                if (0 < c)
                    if (h.getWidth.call(this) > this.content.offsetWidth) console.error(_$_669e[383]), this.options.freezeColumns--;
                    else if (this.merged.cols[c]) console.error(_$_669e[384]), this.options.freezeColumns = 0;
                else
                    for (h.headers.call(this), c = 0; c < this.visibleRows.length; c++) h.update.call(this, this.visibleRows[c])
            };
            h[_$_669e[11]] = function(c) {
                h.applyReset.call(this);
                var b = this.options.freezeColumns;
                this.options.freezeColumns = c;
                this.options.freezeColumnControl &&
                    !c ? this.thead.lastChild.firstChild.classList.add(_$_669e[385]) : this.thead.lastChild.firstChild.classList.remove(_$_669e[385]);
                h.width.call(this);
                J.call(this, !0);
                ca.refresh.call(this);
                Q.call(this.parent, {
                    worksheet: this,
                    action: _$_669e[94],
                    newValue: c,
                    oldValue: b
                });
                L.call(this, _$_669e[94], {
                    column: c
                })
            };
            h.reset = function() {
                h[_$_669e[11]].call(this, 0)
            };
            h.applyReset = function() {
                for (var c, b = 1; b <= this.options.freezeColumns; b++) {
                    if (c = this.thead.lastChild) c.children[b].classList.remove(_$_669e[58]), c.children[b].style.left =
                        _$_669e[3];
                    for (var a = 0; a < this.rows.length; a++)
                        if (c = this.rows[a].element) c.children[b].classList.remove(_$_669e[58]), c.children[b].style.left = _$_669e[3];
                    if (this.options.footers)
                        for (a = 0; a < this.options.footers.length; a++) this.tfoot.children[a].children[b].classList.remove(_$_669e[58]), this.tfoot.children[a].children[b].style.left = _$_669e[3]
                }
            };
            h.getWidth = function(c) {
                c = c ? 0 : Ca.call(this) ? 50 : 0;
                if (0 < this.options.freezeColumns)
                    for (var b = 0; b < this.options.freezeColumns; b++) na.isVisible.call(this, b) && (c += parseInt(this.options.columns[b].width));
                return c
            };
            h.start = function(c, b) {
                c = this.parent.helper;
                h.event = {
                    x: b,
                    d: b
                };
                this.content.scrollLeft = 0;
                b = h.getWidth.call(this);
                c.style.display = _$_669e[24];
                c.style.top = this.content.offsetTop + _$_669e[57];
                c.style.left = b + _$_669e[57];
                c.style.width = _$_669e[200];
                c.style.height = this.content.offsetHeight + _$_669e[57];
                this.thead.lastChild.firstChild.classList.remove(_$_669e[385])
            };
            h.end = function(c) {
                c = h.event;
                null !== c.d && this.setFreezeColumns(c.d);
                this.parent.helper.style.display = _$_669e[3];
                h.event = null
            };
            h.cancel =
                function() {
                    h.event.d = null;
                    h.end.call(this, !1)
                };
            h.move = function(c) {
                var b = c.target.getAttribute(_$_669e[63]);
                if (null != b) {
                    b = parseInt(b);
                    var a = (c = c.target.clientWidth / 2 > c.offsetX) ? b : b + 1;
                    this.merged.cols[a] && 0 < a || (h.event.d = a, b = this.headers[b].getBoundingClientRect(), a = this.content.getBoundingClientRect(), this.parent.helper.style.left = (c ? b.left : b.right) - a.left + _$_669e[57])
                }
            };
            return h
        }(),
        ha = function() {
            var h = function(a, d) {
                    if (this.options[a]) {
                        for (var e = {}, g = Object.keys(this.options[a]), f = 0; f < g.length; f++) e[d &&
                            void 0 != d[g[f]] ? d[g[f]] : g[f]] = this.options[a][g[f]];
                        this.options[a] = e
                    }
                },
                c = function() {
                    if (0 < this.options.minSpareRows) {
                        for (var a = 0, d = this.rows.length - 1; 0 <= d; d--) {
                            for (var e = !1, g = 0; g < this.options.columns.length; g++) R.call(this, g, d) && (e = !0);
                            if (e) break;
                            else a++
                        }
                        0 < this.options.minSpareRows - a && this.insertRow(this.options.minSpareRows - a)
                    }
                    if (0 < this.options.minSpareCols) {
                        a = 0;
                        for (g = this.options.columns.length - 1; 0 <= g; g--) {
                            e = !1;
                            for (d = 0; d < this.rows.length; d++) R.call(this, g, d) && (e = !0);
                            if (e) break;
                            else a++
                        }
                        0 < this.options.minSpareCols -
                            a && this.insertColumn(this.options.minSpareCols - a)
                    }
                },
                b = function() {};
            b.build = function() {
                var a = this;
                this.content = document.createElement(_$_669e[32]);
                this.content.classList.add(_$_669e[386]);
                this.content.addEventListener(_$_669e[387], function(e) {
                    if (!a.edition) return e.preventDefault(), !1
                });
                this.table = document.createElement(_$_669e[388]);
                this.thead = document.createElement(_$_669e[389]);
                this.tbody = document.createElement(_$_669e[390]);
                this.tfoot = document.createElement(_$_669e[391]);
                !1 === this.options.gridline &&
                    this.table.classList.add(_$_669e[392]);
                this.headers = [];
                this.colgroup = [];
                this.colgroupContainer = document.createElement(_$_669e[393]);
                var d = document.createElement(_$_669e[213]);
                d.setAttribute(_$_669e[196], _$_669e[394]);
                this.colgroupContainer.appendChild(d);
                this.headerContainer = document.createElement(_$_669e[183]);
                d = document.createElement(_$_669e[185]);
                d.classList.add(_$_669e[395]);
                if (1 == this.parent.edition) this.options.freezeRowControl && !this.options.freezeRows && d.classList.add(_$_669e[382]), this.options.freezeColumnControl &&
                    !this.options.freezeColumns && d.classList.add(_$_669e[385]);
                else if (this.options.freezeRowControl || this.options.freezeColumnControl) this.options.freezeRowControl = !1, this.options.freezeColumnControl = !1, console.log(_$_669e[396]);
                this.headerContainer.appendChild(d);
                this.thead.appendChild(this.headerContainer);
                this.table.classList.add(_$_669e[397]);
                this.table.setAttribute(_$_669e[398], _$_669e[126]);
                this.table.setAttribute(_$_669e[399], _$_669e[126]);
                this.table.setAttribute(_$_669e[400], _$_669e[401]);
                this.table.appendChild(this.colgroupContainer);
                this.table.appendChild(this.thead);
                this.table.appendChild(this.tbody);
                this.table.appendChild(this.tfoot);
                this.options.defaultColAlign && this.table.classList.add(_$_669e[402] + this.options.defaultColAlign);
                this.options.textOverflow || this.table.classList.add(_$_669e[403]);
                !1 === this.options.selectLockedCells && this.table.classList.add(_$_669e[404]);
                this.corner = document.createElement(_$_669e[32]);
                this.corner.className = _$_669e[405];
                this.corner.setAttribute(_$_669e[400], _$_669e[406]);
                this.corner.setAttribute(_$_669e[407],
                    _$_669e[408]);
                0 == this.options.selectionCopy && (this.corner.style.display = _$_669e[52]);
                this.scroll = J.build.call(a);
                this.content.appendChild(this.scroll);
                this.content.appendChild(this.table);
                this.content.appendChild(this.corner);
                this.element.appendChild(this.content);
                1 == this.options.tableOverflow && (this.content.classList.add(_$_669e[66]), this.options.tableHeight || this.parent.element.style.height || (this.options.tableHeight = 300), this.options.tableWidth || this.parent.element.style.width || (this.options.tableWidth =
                    document.body.offsetWidth - 18), this.options.tableWidth && (this.content.style.width = parseInt(this.options.tableWidth) + _$_669e[57]), this.options.pagination ? this.content.classList.add(_$_669e[409]) : this.options.tableHeight && (this.content.style.height = parseInt(this.options.tableHeight) + _$_669e[57]));
                this.options.resize && (this.content.style.resize = this.options.resize);
                this.options.freezeColumns > this.options.columns.length && console.error(_$_669e[410]);
                sa.build.call(this);
                Ja.build.call(this);
                xa.build.call(this);
                Aa.build.call(this);
                Ea.build.call(this);
                Ka.build.call(this);
                ia.build.call(this);
                la.onload.call(this);
                c.call(this);
                J.call(this)
            };
            b.references = function(a) {
                var d = [],
                    e = [],
                    g = null,
                    f = null,
                    k = null,
                    l = null;
                for (g = 0; g < this.options.columns.length; g++)
                    if (f = this.headers[g]) k = f.getAttribute(_$_669e[63]), k != g && (f.setAttribute(_$_669e[63], g), f.getAttribute(_$_669e[210]) || (f.innerHTML = H.getColumnName(g)));
                for (var n = 0; n < this.rows.length; n++) f = this.rows[n], l = f.y, l != n && (f.y = n, d[l] = n, f.element && (f.element.setAttribute(_$_669e[59],
                    n), f.element.children[0].setAttribute(_$_669e[59], n), g = this.options.rows && this.options.rows[l] && this.options.rows[l].title ? this.options.rows[l].title : n + 1, f.element.children[0].innerHTML = g));
                var p = function(v, w) {
                    f = this.records[w][v];
                    k = f.x;
                    l = f.y;
                    k != v && (f.x = v, f.element && f.element.setAttribute(_$_669e[63], v));
                    l != w && (f.y = w, f.element && f.element.setAttribute(_$_669e[59], w));
                    if (k != v || l != w) {
                        var x = H.getColumnNameFromCoords(k, l);
                        v = H.getColumnNameFromCoords(v, w);
                        e.push([x, v])
                    }
                };
                if (a) {
                    n = function(v, w) {
                        v && v[w] &&
                            delete v[w]
                    };
                    var q = Object.keys(a);
                    for (g = 0; g < q.length; g++) n(this, q[g]), n(this.options.meta, q[g]), n(this.options.cells, q[g]), n(this.options.style, q[g]), n(this.options.comments, q[g])
                }
                for (n = 0; n < this.rows.length; n++)
                    for (g = 0; g < this.options.columns.length; g++) p.call(this, g, n);
                e = Object.fromEntries(e);
                h.call(this, _$_669e[411], d);
                h.call(this, _$_669e[334], e);
                h.call(this, _$_669e[335], e);
                h.call(this, _$_669e[127], e);
                h.call(this, _$_669e[336], e);
                h.call(this, _$_669e[314], e);
                this.options.mergeCells && sa.build.call(this);
                G.call(this.parent, _$_669e[412], this, e, a);
                Y.tracking.call(this, e, a);
                Y.updateAll.call(this, e, a);
                La.updateAll.call(this, e, a);
                this.resetBorders(_$_669e[115], !0);
                G.call(this.parent, _$_669e[413], this, e, a);
                return e
            };
            b.update = function(a) {
                a && c.call(this);
                if (typeof this.parent.config.updateTable == _$_669e[110])
                    for (var d = 0; d < this.rows.length; d++)
                        for (var e = 0; e < this.options.columns.length; e++) a = this.records[d][e], this.parent.config.updateTable.call(this, this, a.element, e, d, R.call(this, e, d), a.r ? a.r : a.v);
                ia.refresh.call(this)
            };
            return b
        }(),
        va = function() {
            var h = {
                spreadsheet: function(c, b) {
                    var a = {
                        name: null,
                        config: {},
                        el: c,
                        element: c,
                        plugins: [],
                        worksheets: [],
                        history: [],
                        historyIndex: -1,
                        queue: [],
                        ignoreEvents: !1,
                        ignoreHistory: !1,
                        ignorePersistence: !1
                    };
                    Ha.spreadsheet.call(a, b);
                    a.config.license && (r.license = a.config.license);
                    c.spreadsheet = a;
                    c.jspreadsheet = c.jexcel = a.worksheets;
                    r.spreadsheet.push(a);
                    a.name || (a.name = jSuites.guid());
                    h.bind(a.config.root ? a.config.root : document);
                    var d = a.config.tabs;
                    typeof d !== _$_669e[109] && (d = {
                        allowCreate: d ?
                            !0 : !1,
                        hideHeaders: d ? !1 : !0,
                        allowChangePosition: a.config.allowMoveWorksheet ? !0 : !1
                    });
                    d.maxWidth || (d.maxWidth = c.offsetWidth - 50 + _$_669e[57]);
                    d.animation = void 0 == d.animation ? !0 : d.animation;
                    d.onbeforecreate = function() {
                        a.createWorksheet();
                        return !1
                    };
                    d.onclick = function(g, f, k, l, n) {
                        0 <= k && a.openWorksheet(k, !0)
                    };
                    d.onchangeposition = function(g, f, k) {
                        a.updateWorksheet(f, k)
                    };
                    jSuites.tabs(c, d);
                    c.classList.add(_$_669e[414]);
                    1 == b.fullscreen && c.classList.add(_$_669e[23]);
                    var e = document.createElement(_$_669e[32]);
                    a.toolbar =
                        document.createElement(_$_669e[32]);
                    a.toolbar.className = _$_669e[415];
                    e.appendChild(a.toolbar);
                    a.filter = la.build.call(a);
                    e.appendChild(a.filter);
                    a.helper = document.createElement(_$_669e[32]);
                    a.helper.className = _$_669e[416];
                    e.appendChild(a.helper);
                    a.loading = document.createElement(_$_669e[32]);
                    a.loading.classList.add(_$_669e[417]);
                    e.appendChild(a.loading);
                    e.appendChild(pa.build(a));
                    e.appendChild(Xa(a));
                    a.textarea = document.createElement(_$_669e[418]);
                    a.textarea.className = _$_669e[419];
                    a.textarea.tabIndex =
                        _$_669e[420];
                    e.appendChild(a.textarea);
                    d.position === _$_669e[181] ? c.insertBefore(e, c.children[0]) : c.insertBefore(e, c.children[1]);
                    c.appendChild(Oa.call(a, r.license));
                    a.createWorksheet = ua.createWorksheet;
                    a.deleteWorksheet = ua.deleteWorksheet;
                    a.renameWorksheet = ua.renameWorksheet;
                    a.updateWorksheet = ua.updateWorksheet;
                    a.openWorksheet = ua.openWorksheet;
                    a.moveWorksheet = ua.moveWorksheet;
                    a.getWorksheet = ua.getWorksheet;
                    a.getWorksheetActive = ua.getWorksheetActive;
                    a.getWorksheetInstance = ua.getWorksheetInstance;
                    a.getConfig =
                        function() {
                            var g = a.config;
                            g.worksheets = [];
                            for (var f = 0; f < a.worksheets.length; f++) g.worksheets.push(a.worksheets[f].getConfig());
                            return g
                        };
                    a.fullscreen = Pa;
                    a.progress = Qa;
                    a.dispatch = G;
                    a.setHistory = Q;
                    a.resetHistory = Q.reset;
                    a.undo = Q.undo;
                    a.redo = Q.redo;
                    a.setToolbar = ma[_$_669e[11]];
                    a.getToolbar = ma[_$_669e[56]];
                    a.showToolbar = ma.show;
                    a.hideToolbar = ma.hide;
                    a.tools = e;
                    La.build.call(a);
                    b.plugins && wa.call(a, b.plugins);
                    a.setPlugins = function(g) {
                        wa.call(a, g)
                    };
                    c = Object.keys(r.extensions);
                    if (c.length)
                        for (d = 0; d < c.length; d++)
                            if (typeof r.extensions[c[d]].oninit ==
                                _$_669e[110]) r.extensions[c[d]].oninit(a, b);
                    return a
                },
                worksheet: function(c, b) {
                    var a = {};
                    c.worksheets.push(a);
                    a.parent = c;
                    Ha(a);
                    Za(a);
                    ka(a);
                    ca(a);
                    pa(a);
                    Ka(a);
                    Ja(a);
                    sa(a);
                    Na(a);
                    ab(a);
                    qa(a);
                    Y(a);
                    U(a);
                    aa(a);
                    na(a);
                    cb(a);
                    bb(a);
                    Fa(a);
                    ia(a);
                    Aa(a);
                    xa(a);
                    la(a);
                    ma(a);
                    za(a);
                    eb(a);
                    fb(a);
                    Ea(a);
                    ua(a);
                    ea(a);
                    ja(a);
                    $a(a);
                    db(a);
                    La(a);
                    Ma(a);
                    Ha.worksheet.call(a, b);
                    !r.license && a.options.license && (r.license = a.options.license);
                    c.element.jexcel = c.element.jexcel ? c.worksheets : a;
                    wa.execute.call(a.parent, _$_669e[421], [a]);
                    a.onload =
                        function() {
                            this.options.worksheetName || (this.options.worksheetName = _$_669e[422] + ua.nextName());
                            this.options.freezeColumns && (this.options.freezeColumns = parseInt(this.options.freezeColumns));
                            this.options.freezeRows && (this.options.freezeRows = parseInt(this.options.freezeRows));
                            this.options.style || (this.options.style = {});
                            var d = this.options.worksheetName;
                            I(d) && console.log(_$_669e[423] + d + _$_669e[424]);
                            I(d, a);
                            this.element = c.element.tabs.appendElement(this.options.worksheetName, function(e, g) {
                                a.options.worksheetState ===
                                    _$_669e[119] && (g.style.display = _$_669e[52])
                            });
                            this.element.classList.add(_$_669e[425]);
                            this.element.jexcel = this;
                            ha.build.call(this);
                            wa.execute.call(this.parent, _$_669e[426], [this]);
                            this.options.data && 500 < this.options.data.length && !J.limited.call(this) && !this.options.pagination && console.error(_$_669e[427]);
                            r.current = this
                        };
                    return a
                },
                row: function(c, b) {
                    this.rows[c] || (this.rows[c] = {
                        element: null,
                        y: c
                    });
                    b && (this.rows[c].id = b, za[_$_669e[11]].call(this, b));
                    (b = this.options.rows) && b[c] && (b[c].height && (this.rows[c].height =
                        parseInt(b[c].height)), !1 === b[c].visible && (this.rows[c].visible = !1));
                    this.records[c] || (this.records[c] = []);
                    this.options.data[c] || (this.options.data[c] = this.dataType ? {} : []);
                    for (b = 0; b < this.options.columns.length; b++) h.cell.call(this, b, c)
                },
                cell: function(c, b, a) {
                    typeof a == _$_669e[21] && (a = R.call(this, c, b), null === a || void 0 === a) && (a = _$_669e[3], R.call(this, c, b, a));
                    this.records[b][c] || (this.records[b][c] = {
                        element: null,
                        x: c,
                        y: b
                    }, Z(a) || (this.records[b][c].v = a))
                },
                bind: function(c) {
                    for (var b = Object.keys(Ga), a = 0; a <
                        b.length; a++)(b[a] == _$_669e[428] ? window : c).addEventListener(b[a], Ga[b[a]], b[a] == _$_669e[429] || b[a] == _$_669e[430] || b[a] == _$_669e[431] ? {
                        passive: !1
                    } : {})
                },
                unbind: function(c) {
                    for (var b = Object.keys(Ga), a = 0; a < b.length; a++)(b[a] == _$_669e[428] ? window : c).removeEventListener(b[a], Ga[b[a]])
                }
            };
            return h
        }(),
        eb = function() {
            var h = function(b) {
                b.getWidth = h[_$_669e[56]];
                b.setWidth = h[_$_669e[11]];
                b.autoWidth = h.auto
            };
            h.auto = function(b) {
                if (b.length) {
                    var a = [],
                        d = [],
                        e;
                    this.table.style.tableLayout = _$_669e[432];
                    for (var g = 0; g <
                        b.length; g++)
                        if (e = this.headers[b[g]]) e = e.offsetWidth, 26 < e && (a.push(b[g]), d.push(e));
                    this.table.style.tableLayout = _$_669e[3];
                    this.setWidth(a, d)
                }
            };
            h[_$_669e[56]] = function(b) {
                var a = this.options.columns;
                if (typeof b === _$_669e[21]) {
                    b = [];
                    for (var d = 0; d < a.length; d++) b.push(a[d].width)
                } else b = a[b].width;
                return b
            };
            var c = function(b, a) {
                this.options.columns[b].width = a;
                this.colgroup[b] && this.colgroup[b].setAttribute(_$_669e[196], a);
                !1 === this.options.columns[b].visible && (this.options.columns[b].visible = !0)
            };
            h[_$_669e[11]] =
                function(b, a, d) {
                    if (!T.call(this.parent, this)) return !1;
                    if (a) {
                        if (Array.isArray(b)) {
                            d || (d = []);
                            for (var e = 0; e < b.length; e++) {
                                d[e] || (d[e] = this.options.columns[b[e]].width);
                                var g = Array.isArray(a) && a[e] ? a[e] : a;
                                c.call(this, b[e], g)
                            }
                        } else d || (d = this.options.columns[b].width), c.call(this, b, a);
                        this.refreshBorders();
                        Q.call(this.parent, {
                            worksheet: this,
                            action: _$_669e[90],
                            column: b,
                            oldValue: d,
                            newValue: a
                        });
                        L.call(this, _$_669e[90], {
                            column: b,
                            width: a,
                            oldWidth: d
                        });
                        G.call(this.parent, _$_669e[433], this, b, a, d);
                        J.call(this,
                            !0)
                    }
                };
            return h
        }(),
        fb = function() {
            var h = function(b) {
                b.getHeight = h[_$_669e[56]];
                b.setHeight = h[_$_669e[11]]
            };
            h[_$_669e[56]] = function(b) {
                if (void 0 === b) {
                    b = [];
                    for (var a = 0; a < this.rows.length; a++) {
                        var d = aa.height.call(this, a);
                        d && (b[a] = d)
                    }
                } else b = aa.height.call(this, b);
                return b
            };
            var c = function(b, a) {
                a = parseInt(a);
                var d = aa.height.call(this, b);
                this.rows[b].element && (this.rows[b].element.style.height = a + _$_669e[57]);
                !1 === this.rows[b].visible && (this.rows[b].visible = !0, this.rows[b].element && (this.rows[b].element.style.display =
                    _$_669e[3]));
                this.options.rows[b] || (this.options.rows[b] = {});
                this.options.rows[b].height = a;
                this.rows[b].height = a;
                return d
            };
            h[_$_669e[11]] = function(b, a, d) {
                if (!T.call(this.parent, this)) return !1;
                var e, g = typeof d == _$_669e[21] ? !0 : !1;
                if (a) {
                    if (Array.isArray(b))
                        for (!0 === g && (d = []), e = 0; e < b.length; e++) {
                            var f = Array.isArray(a) && a[e] ? a[e] : a;
                            c.call(this, b[e], f);
                            1 == g && d.push(g)
                        } else e = c.call(this, b, a), 1 == g && (d = e);
                    this.refreshBorders();
                    Q.call(this.parent, {
                        worksheet: this,
                        action: _$_669e[91],
                        row: b,
                        oldValue: d,
                        newValue: a
                    });
                    L.call(this, _$_669e[91], {
                        row: b,
                        height: a,
                        oldHeight: d
                    });
                    G.call(this.parent, _$_669e[434], this, b, a, d);
                    J.call(this, !0)
                }
            };
            return h
        }(),
        la = function() {
            var h = function(f) {
                    var k = e[f];
                    f = a[e[f]];
                    if (!f.element) {
                        f.element = document.createElement(_$_669e[435]);
                        f.element.innerHTML = f.value;
                        var l = document.createElement(_$_669e[269]);
                        l.type = _$_669e[301];
                        l.value = k;
                        l.o = f;
                        f.element.insertBefore(l, f.element.firstChild)
                    }
                    f.element.firstChild.checked = f.selected;
                    return f.element
                },
                c = null,
                b = null,
                a = [],
                d = 0,
                e = null,
                g = function(f) {
                    f.setFilter =
                        g[_$_669e[11]];
                    f.getFilter = g[_$_669e[56]];
                    f.openFilter = g.open;
                    f.closeFilter = g.close;
                    f.resetFilters = g.reset;
                    f.showFilter = g.show;
                    f.hideFilter = g.hide
                };
            g.isVisible = function() {
                return b
            };
            g[_$_669e[11]] = function(f, k) {
                var l = this.headers[f];
                f = this.options.columns[f];
                Array.isArray(k) && 0 < k.length ? (l && l.classList.add(_$_669e[436]), f.filters = k) : (l && l.classList.remove(_$_669e[436]), f.filters = null);
                g.update.call(this)
            };
            g[_$_669e[56]] = function(f) {
                if (f) return this.options.columns[f].filters || null;
                f = {};
                for (var k = 0; k <
                    this.options.columns.length; k++) f[k] = this.options.columns[k].filters || null;
                return f
            };
            g.open = function(f) {
                c = parseInt(f);
                if (this.headers[c].classList.contains(_$_669e[215])) {
                    b = !0;
                    var k = this.parent.element.getBoundingClientRect(),
                        l = this.headers[c].getBoundingClientRect();
                    f = l.left - k.left;
                    var n = l.top - k.top + l.height;
                    l = this.parent.filter;
                    k = k.left + k.width;
                    l.style.display = _$_669e[24];
                    f + l.offsetWidth > k && (f = k - (l.offsetWidth + 18));
                    l.style.top = n + _$_669e[57];
                    l.style.left = f + _$_669e[57];
                    l.children[0].focus();
                    f = l.children[1].selectAll;
                    l.children[1].textContent = _$_669e[3];
                    l.children[1].appendChild(f);
                    f.children[0].checked = this.options.columns[c].filters ? !1 : !0;
                    f = xa.execute.call(this, null, c);
                    f = H.invert(f);
                    l = [];
                    for (k = 0; k < this.rows.length; k++)
                        if (f[k]) {
                            n = _$_669e[3] + R.call(this, c, k);
                            var p = _$_669e[3] + qa.processed.call(this, c, k, !0);
                            n.substr(0, 1) == _$_669e[18] && (n = p);
                            l.push([n, p])
                        } l.sort(function(q, v) {
                        return q[0] > v[0] ? 1 : q[0] < v[0] ? -1 : 0
                    });
                    a = [];
                    for (k = 0; k < l.length; k++) a[l[k][1]] = {
                        value: l[k][1],
                        element: null,
                        key: l[k][0]
                    };
                    (f = G.call(this.parent,
                        _$_669e[438], this, c, a)) && (a = f);
                    typeof a[_$_669e[3]] !== _$_669e[21] && (delete a[_$_669e[3]], a[_$_669e[3]] = {
                        value: _$_669e[259] + D(_$_669e[439]) + _$_669e[281],
                        element: null
                    });
                    g.search.call(this, _$_669e[3])
                } else console.log(_$_669e[437])
            };
            g.close = function(f) {
                if (f) {
                    f = this.parent.filter.children[1];
                    if (1 == f.selectAll.firstChild.checked && e.length == Object.keys(a).length) f = null;
                    else {
                        f = 1 == f.currentSelection.firstChild.checked ? this.options.columns[c].filters || [] : [];
                        for (var k = 0; k < e.length; k++) a[e[k]].selected && f.push(e[k])
                    }
                    g[_$_669e[11]].call(this,
                        c, f)
                }
                this.parent.filter.style.display = _$_669e[3];
                this.parent.filter.children[0].value = _$_669e[3];
                c = null;
                b = !1;
                a = null;
                d = 0;
                e = null
            };
            g.reset = function(f) {
                g.updateDOM.call(this, f, !1);
                Ba.reset.call(this);
                g.update.call(this)
            };
            g.execute = function(f) {
                for (var k = this.options.columns, l = [], n = 0; n < this.options.data.length; n++) {
                    for (var p = !0, q = 0; q < k.length; q++)
                        if (k[q].filters && f !== q) {
                            var v = _$_669e[3] + R.call(this, q, n),
                                w = _$_669e[3] + this.getLabelFromCoords(q, n, !0); - 1 == k[q].filters.indexOf(v) && -1 == k[q].filters.indexOf(w) &&
                                (p = !1)
                        } p && l.push(n)
                }
                return l
            };
            g.update = function() {
                Ba.reset.call(this);
                this.resetSelection();
                this.resetBorders();
                var f = xa.execute.call(this, null);
                if (typeof this.parent.config.onbeforefilter == _$_669e[110]) {
                    var k = g[_$_669e[56]].call(this),
                        l = G.call(this.parent, _$_669e[440], this, k, f);
                    if (l) f = l;
                    else if (!1 === l) return !1
                }
                this.pageNumber = 0;
                f && f.length < this.rows.length && Ba[_$_669e[11]].call(this, f);
                ia.refresh.call(this);
                J.call(this, !0);
                G.call(this.parent, _$_669e[441], this, k, f)
            };
            g.show = function(f) {
                g.updateDOM.call(this,
                    f, !0)
            };
            g.hide = function(f) {
                g.updateDOM.call(this, f, !1, !0);
                g.update.call(this)
            };
            g.search = function(f) {
                for (var k = this.parent.filter.children[1]; k.children[1];) k.removeChild(k.lastChild);
                k.currentSelection.firstChild.checked = !1;
                if (f) {
                    e = [];
                    f = new RegExp(f, _$_669e[442]);
                    for (var l = Object.keys(a), n = 0; n < l.length; n++) a[l[n]] && (_$_669e[3] + a[l[n]].value).match(f) ? (e.push(l[n]), a[l[n]].selected = !0) : a[l[n]].selected = !1;
                    e.length && this.options.columns[c].filters && (k.firstChild.checked = !0, k.appendChild(k.currentSelection))
                } else {
                    e =
                        Object.keys(a);
                    if (f = this.options.columns[c].filters)
                        for (n = 0; n < e.length; n++) a[e[n]].selected = !1, 0 <= f.indexOf(e[n]) && (a[e[n]].selected = !0);
                    else
                        for (n = 0; n < e.length; n++) a[e[n]].selected = !0;
                    k.firstChild.checked = f ? !1 : !0
                }
                for (n = d = 0; n < e.length; n++) 200 > n && (f = d++, k.appendChild(h(f)))
            };
            g.build = function() {
                var f = document.createElement(_$_669e[32]);
                f.className = _$_669e[443];
                f.onclick = function(p) {
                    p.target.tagName == _$_669e[444] && p.target.o && (p.target.o.selected = p.target.checked);
                    p = !0;
                    for (var q = 1; q < this.children.length; q++) this.children[q].children[0].checked ||
                        (p = !1);
                    this.children[0].children[0].checked = p
                };
                var k = document.createElement(_$_669e[269]);
                k.type = _$_669e[69];
                k.placeholder = D(_$_669e[445]);
                k.classList.add(_$_669e[446]);
                k.classList.add(_$_669e[270]);
                var l = document.createElement(_$_669e[269]);
                l.type = _$_669e[447];
                l.value = _$_669e[448];
                l.className = _$_669e[449];
                var n = document.createElement(_$_669e[32]);
                n.className = _$_669e[450];
                n.appendChild(k);
                n.appendChild(f);
                n.appendChild(l);
                f.currentSelection = document.createElement(_$_669e[435]);
                f.currentSelection.innerHTML =
                    _$_669e[451] + D(_$_669e[452]);
                f.selectAll = document.createElement(_$_669e[435]);
                f.selectAll.innerHTML = _$_669e[453] + D(_$_669e[454]) + _$_669e[281];
                f.selectAll.onclick = function() {
                    for (var p = this.children[0].checked, q = 1; q < this.parentNode.children.length; q++) this.parentNode.children[q].children[0].checked = p;
                    if (1 == p)
                        for (q = 0; q < e.length; q++) a[e[q]].selected = !0;
                    else
                        for (p = Object.keys(a), q = 0; q < p.length; q++) a[p[q]].selected = !1
                };
                jSuites.lazyLoading(f, {
                    loadDown: function(p) {
                        for (p = 0; d < e.length - 1 && 10 > p;) {
                            var q = d++;
                            f.appendChild(h(q));
                            p++
                        }
                    }
                });
                return n
            };
            g.updateDOM = function(f, k, l) {
                var n = this.headers,
                    p = function(q) {
                        1 == k ? n[q].classList.add(_$_669e[215]) : (this.options.columns[q].filters = null, n[q].classList.remove(_$_669e[436]), 1 == l && n[q].classList.remove(_$_669e[215]))
                    };
                if (f) p.call(this, f);
                else
                    for (f = 0; f < n.length; f++) p.call(this, f)
            };
            g.onload = function() {
                var f, k = !1;
                if (f = this.options.columns)
                    for (var l = 0; l < f.length; l++) Array.isArray(f[l].filters) && 0 < f[l].filters.length && this.headers[l] && (this.headers[l].classList.add(_$_669e[436]), k = !0);
                1 == k && g.update.call(this)
            };
            return g
        }(),
        Ba = function() {
            var h = {};
            h[_$_669e[11]] = function(c) {
                var b = null,
                    a = null,
                    d = null,
                    e = [];
                if (this.merged.rows.length) {
                    for (a = 0; a < c.length; a++)
                        if (-1 === e.indexOf(c[a]) && (e.push(c[a]), this.merged.rows[c[a]] && (d = sa.getAffected.call(this.merged.rows, c[a]))))
                            for (b = 0; b < d.length; b++) - 1 === e.indexOf(d[b]) && e.push(d[b]);
                    e.sort(function(g, f) {
                        return g - f
                    })
                } else e = c;
                for (a = 0; a < e.length; a++) this.rows[e[a]].results = !0;
                this.results = e;
                Y.evaluate.call(this)
            };
            h.reset = function() {
                for (var c = 0; c <
                    this.rows.length; c++) !0 === this.rows[c].results && delete this.rows[c].results;
                this.results = null
            };
            h.refresh = function() {
                for (var c = [], b = 0; b < this.rows.length; b++) !0 === this.rows[b].results && c.push(b);
                return c
            };
            return h
        }(),
        xa = function() {
            var h = function(e) {
                    var g = document.createElement(_$_669e[32]);
                    if (0 < e.options.pagination && e.options.paginationOptions && 0 < e.options.paginationOptions.length) {
                        e.paginationDropdown = document.createElement(_$_669e[455]);
                        e.paginationDropdown.classList.add(_$_669e[456]);
                        e.paginationDropdown.onchange =
                            function() {
                                e.options.pagination = parseInt(this.value);
                                e.page(0)
                            };
                        for (var f = 0; f < e.options.paginationOptions.length; f++) {
                            var k = document.createElement(_$_669e[457]);
                            k.value = e.options.paginationOptions[f];
                            k.innerHTML = e.options.paginationOptions[f];
                            e.paginationDropdown.appendChild(k)
                        }
                        e.paginationDropdown.value = e.options.pagination;
                        g.appendChild(document.createTextNode(D(_$_669e[458])));
                        g.appendChild(e.paginationDropdown);
                        g.appendChild(document.createTextNode(D(_$_669e[459])))
                    }
                    return g
                },
                c = function(e, g) {
                    var f =
                        this.merged.rows;
                    if (f[g]) {
                        for (; 0 < g && f[g - 1];) g--;
                        for (; f[g];) e.push(g++)
                    } else e.push(g)
                },
                b = function(e) {
                    var g = this.searchInput.parentNode;
                    e ? g.classList.add(_$_669e[461]) : setTimeout(function() {
                        g.classList.remove(_$_669e[461])
                    }, 500)
                },
                a = function(e, g) {
                    for (var f, k = 0; k < this.options.columns.length; k++) {
                        f = _$_669e[3] + R.call(this, k, e);
                        if (f.match(g)) return !0;
                        f = _$_669e[3] + this.getLabelFromCoords(k, e);
                        if (f.match(g)) return !0
                    }
                    return !1
                },
                d = function(e) {
                    e.search = d.update;
                    e.resetSearch = d.reset;
                    e.updateSearch = d.refresh;
                    e.showSearch = d.show;
                    e.hideSearch = d.hide
                };
            d.execute = function(e, g) {
                null === e ? e = this.searchInput.value || _$_669e[3] : e !== this.searchInput.value && (this.searchInput.value = e);
                this.pageNumber = 0;
                var f = [],
                    k = null,
                    l = null,
                    n = this.options.freezeRows || 0;
                g = la.execute.call(this, g);
                if (e) try {
                    var p = new RegExp(e, _$_669e[442]);
                    if (g && g.length) {
                        var q = {};
                        for (k = 0; k < g.length; k++) q[g[k]] = !0
                    } else q = null;
                    for (k = 0; k < this.rows.length; k++) l = !1, (!q || q[k]) && (l = this.parent.config.onsearchrow ? G.call(this.parent, _$_669e[462], this, k, e) : a.call(this,
                        k, p)) && c.call(this, f, k)
                } catch (v) {} else f = g;
                if (n)
                    for (k = 0; k < n; k++) f.push(k);
                if (f.length) {
                    f = new Set(f);
                    f = Array.from(f);
                    for (k = 0; k < this.rows.length; k++) !0 === this.rows[k].results && f.push(k);
                    f.sort(function(v, w) {
                        return v - w
                    })
                }
                return f
            };
            d.update = function(e) {
                b.call(this, !0);
                Ba.reset.call(this);
                this.resetSelection();
                this.resetBorders();
                G.call(this.parent, _$_669e[463], this, e);
                var g = d.execute.call(this, e);
                if (typeof this.parent.config.onbeforesearch == _$_669e[110]) {
                    var f = G.call(this.parent, _$_669e[464], this, e, g);
                    if (f) g = f;
                    else if (!1 === f) return b.call(this, !1), !1
                }
                this.pageNumber = 0;
                g && g.length < this.rows.length && Ba[_$_669e[11]].call(this, g);
                ia.refresh.call(this);
                J.call(this, !0);
                G.call(this.parent, _$_669e[465], this, e, g);
                b.call(this, !1)
            };
            d.refresh = function() {
                J.update.call(this);
                ia.refresh.call(this)
            };
            d.reset = function() {
                this.searchInput.value = _$_669e[3];
                Ba.reset.call(this);
                J.update.call(this);
                ia.refresh.call(this)
            };
            d.show = function() {
                this.options.search = !0;
                this.searchContainer.style.display = _$_669e[3]
            };
            d.hide = function() {
                this.options.search = !1;
                this.searchContainer.style.display = _$_669e[52]
            };
            d.build = function() {
                this.searchContainer = document.createElement(_$_669e[32]);
                this.searchContainer.classList.add(_$_669e[466]);
                this.searchContainer.appendChild(h(this));
                var e = this.searchContainer,
                    g = e.appendChild,
                    f = document.createElement(_$_669e[32]);
                f.innerHTML = D(_$_669e[445]) + _$_669e[144];
                this.searchInput = document.createElement(_$_669e[269]);
                this.searchInput.classList.add(_$_669e[460]);
                f.appendChild(this.searchInput);
                g.call(e, f);
                0 == this.options.search &&
                    (this.searchContainer.style.display = _$_669e[52]);
                this.element.insertBefore(this.searchContainer, this.element.firstChild)
            };
            return d
        }(),
        Xa = function() {
            var h = function(c) {
                c.contextmenu = document.createElement(_$_669e[32]);
                c.contextmenu.className = _$_669e[467];
                jSuites.contextmenu(c.contextmenu, {
                    onclick: function(b, a, d) {
                        b.close()
                    }
                });
                return c.contextmenu
            };
            h.open = function(c, b) {
                if (b[0] === _$_669e[468]) var a = b[1],
                    d = b[2];
                else b[0] === _$_669e[469] ? (a = null, d = b[1]) : b[0] === _$_669e[470] && (a = b[1], d = null);
                var e = h[_$_669e[56]](this,
                    a, d, c, b[0], b[1], b[2]);
                if (typeof this.parent.config.contextMenu == _$_669e[110]) {
                    var g = this.parent.config.contextMenu(this, a, d, c, e, b[0], b[1], b[2]);
                    if (g) e = g;
                    else if (!1 === g) return !1
                }
                c.preventDefault();
                g = wa.execute.call(this.parent, _$_669e[111], [this, a, d, c, e, b[0], b[1], b[2]]);
                e = g[4];
                this.parent.contextmenu.contextmenu.open(c, e)
            };
            h.isOpened = function() {
                return this.parent.contextmenu.contextmenu.isOpened()
            };
            h[_$_669e[56]] = function(c, b, a, d, e, g, f) {
                var k = [],
                    l = c.parent.config,
                    n = -1 != navigator.userAgent.indexOf(_$_669e[471]) ?
                    _$_669e[472] : _$_669e[473];
                e == _$_669e[474] && (1 == l.allowRenameWorksheet && k.push({
                    title: D(_$_669e[475]),
                    onclick: function() {
                        var q = prompt(D(_$_669e[475]), d.target.innerText);
                        q && c.parent.renameWorksheet(g, q)
                    }
                }), 1 == l.allowDeleteWorksheet && k.push({
                    title: D(_$_669e[476]),
                    onclick: function() {
                        confirm(D(_$_669e[477]), d.target.innerText) && c.parent.deleteWorksheet(g)
                    }
                }), k.push({
                    type: _$_669e[478]
                }));
                e == _$_669e[479] && (k.push({
                    title: D(_$_669e[480]),
                    onclick: function() {
                        var q = prompt(D(_$_669e[480]), d.target.innerText);
                        c.setNestedCell(g, f, {
                            title: q
                        })
                    }
                }), k.push({
                    type: _$_669e[478]
                }));
                if (e == _$_669e[470] || e == _$_669e[469] || e == _$_669e[468] || e == _$_669e[479]) k.push({
                    title: D(_$_669e[481]),
                    icon: _$_669e[482],
                    shortcut: n + _$_669e[483],
                    onclick: function() {
                        c.cut()
                    }
                }), k.push({
                    title: D(_$_669e[484]),
                    icon: _$_669e[485],
                    shortcut: n + _$_669e[486],
                    onclick: function() {
                        c.copy()
                    }
                }), navigator && navigator.clipboard && navigator.clipboard.readText && k.push({
                    title: D(_$_669e[487]),
                    icon: _$_669e[488],
                    shortcut: n + _$_669e[489],
                    onclick: function() {
                        c.selectedCell &&
                            navigator.clipboard.readText().then(function(q) {
                                q && c.paste(Math.min(c.selectedCell[0], c.selectedCell[2]), Math.min(c.selectedCell[1], c.selectedCell[3]), q)
                            })
                    }
                }), k.push({
                    type: _$_669e[478]
                });
                e == _$_669e[470] && (1 == c.options.allowInsertColumn && k.push({
                        title: D(_$_669e[490]),
                        onclick: function() {
                            c.insertColumn(1, parseInt(g), 1)
                        }
                    }), 1 == c.options.allowInsertColumn && k.push({
                        title: D(_$_669e[491]),
                        onclick: function() {
                            c.insertColumn(1, parseInt(g), 0)
                        }
                    }), 1 == c.options.allowDeleteColumn && k.push({
                        title: D(_$_669e[492]),
                        onclick: function() {
                            c.deleteColumn(c.getSelectedColumns().length ?
                                void 0 : parseInt(b))
                        }
                    }), 1 == c.options.allowRenameColumn && k.push({
                        title: D(_$_669e[493]),
                        onclick: function() {
                            c.setHeader(g)
                        }
                    }), c.options.data.length || (k.push({
                        type: _$_669e[478]
                    }), k.push({
                        title: D(_$_669e[494]),
                        onclick: function() {
                            c.insertRow(0)
                        }
                    })), 1 == c.options.columnSorting && (k.push({
                        type: _$_669e[478]
                    }), k.push({
                        title: D(_$_669e[495]),
                        onclick: function() {
                            c.orderBy(g, 0)
                        }
                    }), k.push({
                        title: D(_$_669e[496]),
                        onclick: function() {
                            c.orderBy(g, 1)
                        }
                    }), k.push({
                        type: _$_669e[478]
                    })), k.push({
                        title: D(_$_669e[497]),
                        onclick: function() {
                            c.hideColumn(c.getSelectedColumns())
                        }
                    }),
                    k.push({
                        title: D(_$_669e[458]),
                        onclick: function() {
                            c.showColumn(c.getSelectedColumns())
                        }
                    }), k.push({
                        type: _$_669e[478]
                    }));
                if (e == _$_669e[468] || e == _$_669e[469]) 1 == c.options.allowInsertRow && (k.push({
                    title: D(_$_669e[498]),
                    onclick: function() {
                        c.insertRow(1, parseInt(a), 1)
                    }
                }), k.push({
                    title: D(_$_669e[499]),
                    onclick: function() {
                        c.insertRow(1, parseInt(a))
                    }
                })), 1 == c.options.allowDeleteRow && k.push({
                    title: D(_$_669e[500]),
                    onclick: function() {
                        c.deleteRow(c.getSelectedRows().length ? void 0 : parseInt(a))
                    }
                }), k.push({
                    type: _$_669e[478]
                });
                e == _$_669e[469] && (k.push({
                    title: D(_$_669e[497]),
                    onclick: function() {
                        c.hideRow(c.getSelectedRows(!0))
                    }
                }), k.push({
                    title: D(_$_669e[458]),
                    onclick: function() {
                        c.showRow(c.getSelectedRows(!0))
                    }
                }), k.push({
                    type: _$_669e[478]
                }));
                if (e == _$_669e[468] && 1 == c.options.allowComments) {
                    var p = c.records[f][g].element.getAttribute(_$_669e[315]) || _$_669e[3];
                    k.push({
                        title: p ? D(_$_669e[501]) : D(_$_669e[502]),
                        icon: _$_669e[503],
                        onclick: function() {
                            var q = H.getColumnNameFromCoords(g, f),
                                v = prompt(D(_$_669e[504]), p);
                            v && c.setComments(q,
                                v)
                        }
                    });
                    p && k.push({
                        title: D(_$_669e[505]),
                        onclick: function() {
                            var q = H.getColumnNameFromCoords(g, f);
                            c.setComments(q, _$_669e[3])
                        }
                    });
                    k.push({
                        type: _$_669e[478]
                    })
                }
                e == _$_669e[506] && (k.push({
                    title: D(_$_669e[507]),
                    onclick: function() {
                        c.showRow(c.getSelectedRows(!0));
                        c.showColumn(c.getSelectedColumns())
                    }
                }), k.push({
                    type: _$_669e[478]
                }));
                l.allowExport && k.push({
                    title: D(_$_669e[508]),
                    icon: _$_669e[509],
                    shortcut: n + _$_669e[510],
                    onclick: function() {
                        c.download()
                    }
                });
                l.about && k.push({
                    title: D(_$_669e[511]),
                    icon: _$_669e[512],
                    onclick: function() {
                        !0 === l.about ? alert(M().print()) : alert(l.about)
                    }
                });
                return k
            };
            return h
        }(),
        ma = function() {
            var h = function(b) {
                b.showToolbar = function() {
                    h.show.call(b.parent)
                };
                b.hideToolbar = function() {
                    h.hide.call(b.parent)
                };
                b.refreshToolbar = function() {
                    h.update.call(b.parent, b, b.selectedCell)
                }
            };
            h[_$_669e[11]] = function(b) {
                this.toolbar.innerHTML = _$_669e[3];
                this.config.toolbar = h[_$_669e[56]].call(this, b);
                h.show.call(this)
            };
            h[_$_669e[56]] = function(b) {
                b || (b = this.config.toolbar);
                var a = h.getDefault(b);
                b ? Array.isArray(b) ?
                    b = {
                        items: b
                    } : typeof b === _$_669e[110] && (b = b({
                        items: a
                    })) : b = {};
                typeof b !== _$_669e[109] && (b = {});
                b.items || (b.items = a);
                typeof b.responsive == _$_669e[21] && (b.responsive = !0, b.bottom = !1, b.maxWidth = this.element.offsetWidth);
                return b
            };
            h.show = function() {
                this.toolbar.innerHTML || h.build.call(this);
                this.toolbar.style.display = _$_669e[3];
                this.element.classList.add(_$_669e[513])
            };
            h.hide = function() {
                this.toolbar.style.display = _$_669e[52];
                this.element.classList.remove(_$_669e[513])
            };
            h.update = function(b, a) {
                var d = this.toolbarInstance;
                d && (d.options.responsive && (a = this.config.fullscreen ? b.element.offsetWidth : b.options.tableWidth ? parseInt(b.options.tableWidth) : null, d.options.maxWidth = a, d.refresh()), d.update(b))
            };
            h.build = function() {
                for (var b = h[_$_669e[56]].call(this), a = 0; a < b.items.length; a++)
                    if (b.items[a].type == _$_669e[455]) {
                        if (!b.items[a].options && b.items[a].v) {
                            b.items[a].options = [];
                            for (var d = 0; d < b.items[a].v.length; d++) b.items[a].options.push(b.items[a].v[d]);
                            b.items[a].k && (b.items[a].onchange = function(e, g, f, k) {
                                r.current.setStyle(r.current.getSelected(!1,
                                    !0), f.k, k)
                            })
                        }
                    } else b.items[a].type == _$_669e[305] ? (b.items[a].type = _$_669e[191], b.items[a].onclick = function(e, g, f) {
                        f.color || (jSuites.color(f, {
                            onchange: function(k, l) {
                                (k = r.current.getSelected(!1, !0)) && r.current.setStyle(k, f.k, l)
                            },
                            onopen: function(k) {
                                var l;
                                r.current.cursor ? (l = r.current.cursor.style[f.k]) ? k.color.select(k.color.toHex(l)) : k.color.select(_$_669e[3]) : k.color.select(_$_669e[3])
                            }
                        }), f.color.open())
                    }) : !b.items[a].onclick && b.items[a].k && (b.items[a].onclick = function() {
                        var e = r.current.getSelected(!1,
                            !0);
                        e && e.length && r.current.setStyle(e, this.k, this.v)
                    });
                (a = wa.execute.call(this, _$_669e[112], [b])) && (b = a[0]);
                this.toolbarInstance = jSuites.toolbar(this.toolbar, b);
                this.toolbarInstance.application = this;
                b = this.element.tabs.getActive();
                this.toolbarInstance.update(this.worksheets[b])
            };
            h.item = function(b) {
                this.isEditable() ? b.classList.remove(_$_669e[514]) : b.classList.add(_$_669e[514])
            };
            var c = function(b) {
                return r.current && r.current.cursor && r.current.cursor.style[b] ? r.current.cursor.style[b] : !1
            };
            h.getDefault =
                function(b) {
                    var a = [];
                    a.push({
                        content: _$_669e[515],
                        onclick: function() {
                            r.current && r.current.undo()
                        }
                    });
                    a.push({
                        content: _$_669e[516],
                        onclick: function() {
                            r.current && r.current.redo()
                        }
                    });
                    a.push({
                        content: _$_669e[509],
                        onclick: function() {
                            r.current && r.current.download()
                        }
                    });
                    a.push({
                        type: _$_669e[517]
                    });
                    a.push({
                        type: _$_669e[455],
                        width: _$_669e[518],
                        options: [_$_669e[519], _$_669e[520], _$_669e[521], _$_669e[522]],
                        render: function(e) {
                            return _$_669e[523] + e + _$_669e[276] + e + _$_669e[524]
                        },
                        onchange: function(e, g, f, k, l) {
                            r.current &&
                                (l && l != _$_669e[126] ? r.current.setStyle(r.current.getSelected(!1, !0), _$_669e[525], k, !0) : r.current.setStyle(r.current.getSelected(!1, !0), _$_669e[525], _$_669e[3], !0))
                        },
                        updateState: function(e, g, f, k) {
                            if ((e = k.selectedCell) && e.length && k.records[e[1]][e[0]].element) {
                                var l = k.records[e[1]][e[0]].element.style.fontFamily || 0;
                                l && (l = l.replace(/"/g, _$_669e[3]), l = this.picker.options.data.findIndex(function(n) {
                                    return n.toLowerCase() === l.toLowerCase()
                                }), -1 == l && (l = 0));
                                f.picker.setValue(l)
                            }
                            h.item.call(k, f)
                        }
                    });
                    a.push({
                        type: _$_669e[455],
                        width: _$_669e[526],
                        content: _$_669e[527],
                        options: [_$_669e[528], _$_669e[529], _$_669e[530], _$_669e[531], _$_669e[532]],
                        render: function(e) {
                            return _$_669e[533] + e + _$_669e[276] + e + _$_669e[524]
                        },
                        onchange: function(e, g, f, k) {
                            r.current && r.current.setStyle(r.current.getSelected(!1, !0), _$_669e[534], k, !0)
                        },
                        updateState: function(e, g, f, k) {
                            h.item.call(k, f)
                        }
                    });
                    a.push({
                        type: _$_669e[455],
                        options: [_$_669e[180], _$_669e[228], _$_669e[33], _$_669e[535]],
                        render: function(e) {
                            return _$_669e[536] + e + _$_669e[537]
                        },
                        onchange: function(e,
                            g, f, k) {
                            r.current && r.current.setStyle(r.current.getSelected(!1, !0), _$_669e[538], k, !0)
                        },
                        updateState: function(e, g, f, k) {
                            e = 1;
                            k.cursor && (e = window.getComputedStyle(k.cursor), (e = e.getPropertyValue(_$_669e[538])) || (e = k.cursor.style.textAlign || r.current.options.defaultColAlign || _$_669e[228]), e = f.picker.options.data.indexOf(e), 0 > e && (e = 1));
                            f.picker.setValue(e);
                            h.item.call(k, f)
                        }
                    });
                    a.push({
                        content: _$_669e[539],
                        onclick: function(e, g, f) {
                            e = c(_$_669e[540]);
                            e = e === _$_669e[541] || 300 < e ? _$_669e[3] : _$_669e[541];
                            r.current.setStyle(r.current.getSelected(!1,
                                !0), _$_669e[542], e, !0);
                            e === _$_669e[541] ? f.classList.add(_$_669e[543]) : f.classList.remove(_$_669e[543])
                        },
                        updateState: function(e, g, f, k) {
                            h.item.call(k, f);
                            e = c(_$_669e[540]);
                            e === _$_669e[541] || 300 < e ? f.classList.add(_$_669e[543]) : f.classList.remove(_$_669e[543])
                        }
                    });
                    a.push({
                        type: _$_669e[305],
                        content: _$_669e[544],
                        k: _$_669e[305],
                        updateState: function(e, g, f, k) {
                            h.item.call(k, f)
                        }
                    });
                    a.push({
                        type: _$_669e[305],
                        content: _$_669e[545],
                        k: _$_669e[546],
                        updateState: function(e, g, f, k) {
                            h.item.call(k, f)
                        }
                    });
                    if (b === _$_669e[547] ||
                        typeof b === _$_669e[110]) {
                        var d = [_$_669e[182], _$_669e[548], _$_669e[181]];
                        a.push({
                            type: _$_669e[455],
                            options: [_$_669e[549], _$_669e[550], _$_669e[551]],
                            render: function(e) {
                                return _$_669e[552] + e + _$_669e[537]
                            },
                            value: 1,
                            onchange: function(e, g, f, k, l) {
                                r.current && r.current.setStyle(r.current.getSelected(!1, !0), _$_669e[553], d[l], !0)
                            },
                            updateState: function(e, g, f, k) {
                                (e = k.selectedCell) && e.length && k.records[e[1]][e[0]].element && (e = d.indexOf(k.records[e[1]][e[0]].element.style.verticalAlign || _$_669e[548]), f.picker.setValue(e));
                                h.item.call(k, f)
                            }
                        });
                        a.push({
                            content: _$_669e[554],
                            onclick: function(e, g, f) {
                                r.current && (e = r.current.getHighlighted()) && e.length && (g = r.helpers.getColumnNameFromCoords(e[0], e[1]), r.current.records[e[1]][e[0]].merged ? r.current.removeMerge(g) : r.current.setMerge(g, e[2] - e[0] + 1, e[3] - e[1] + 1))
                            },
                            tooltip: D(_$_669e[555]),
                            updateState: function(e, g, f, k) {
                                h.item.call(k, f)
                            }
                        })
                    }
                    a.push({
                        type: _$_669e[455],
                        data: [_$_669e[556], _$_669e[557], _$_669e[558], _$_669e[559], _$_669e[560], _$_669e[561], _$_669e[562], _$_669e[563], _$_669e[564],
                            _$_669e[565]
                        ],
                        columns: 5,
                        render: function(e) {
                            return _$_669e[552] + e + _$_669e[537]
                        },
                        right: !0,
                        onchange: function(e, g, f, k, l) {
                            if (e = r.current.getHighlighted()) {
                                f = g.thickness || 1;
                                g = g.color || _$_669e[566];
                                l = {};
                                for (var n = e[0], p = e[1], q = e[2], v = e[3], w = e[1]; w <= e[3]; w++)
                                    for (var x = e[0]; x <= e[2]; x++) {
                                        var y = H.getColumnNameFromCoords(x, w);
                                        l[y] || (l[y] = _$_669e[3]);
                                        l[y] = k != _$_669e[561] && k != _$_669e[557] || x != n ? (k == _$_669e[558] || k == _$_669e[560]) && x > n ? l[y] + (_$_669e[567] + f + _$_669e[568] + g + _$_669e[353]) : k == _$_669e[556] ? l[y] + (_$_669e[567] +
                                            f + _$_669e[568] + g + _$_669e[353]) : l[y] + _$_669e[569] : l[y] + (_$_669e[567] + f + _$_669e[568] + g + _$_669e[353]);
                                        l[y] = k != _$_669e[556] && k != _$_669e[563] && k != _$_669e[557] || x != q ? l[y] + _$_669e[571] : l[y] + (_$_669e[570] + f + _$_669e[568] + g + _$_669e[353]);
                                        l[y] = k != _$_669e[562] && k != _$_669e[557] || w != p ? (k == _$_669e[558] || k == _$_669e[559]) && w > p ? l[y] + (_$_669e[572] + f + _$_669e[568] + g + _$_669e[353]) : k == _$_669e[556] ? l[y] + (_$_669e[572] + f + _$_669e[568] + g + _$_669e[353]) : l[y] + _$_669e[573] : l[y] + (_$_669e[572] + f + _$_669e[568] + g + _$_669e[353]);
                                        l[y] =
                                            k != _$_669e[556] && k != _$_669e[564] && k != _$_669e[557] || w != v ? l[y] + _$_669e[575] : l[y] + (_$_669e[574] + f + _$_669e[568] + g + _$_669e[353])
                                    }
                                Object.keys(l) && r.current.setStyle(l, null, null, !0)
                            }
                        },
                        onload: function(e, g) {
                            var f = document.createElement(_$_669e[32]),
                                k = document.createElement(_$_669e[32]);
                            f.appendChild(k);
                            var l = jSuites.color(k, {
                                closeOnChange: !1,
                                onchange: function(n, p) {
                                    n.parentNode.children[1].style.color = p;
                                    g.color = p
                                }
                            });
                            k = document.createElement(_$_669e[191]);
                            k.classList.add(_$_669e[192]);
                            k.innerHTML = _$_669e[576];
                            k.onclick = function() {
                                l.open()
                            };
                            f.appendChild(k);
                            e.children[1].appendChild(f);
                            k = document.createElement(_$_669e[32]);
                            jSuites.picker(k, {
                                type: _$_669e[455],
                                data: [1, 2, 3, 4, 5],
                                render: function(n) {
                                    return _$_669e[577] + n + _$_669e[578]
                                },
                                onchange: function(n, p, q, v) {
                                    g.thickness = v
                                },
                                width: _$_669e[579]
                            });
                            e.children[1].appendChild(k);
                            k = document.createElement(_$_669e[32]);
                            k.style.flex = _$_669e[580];
                            e.children[1].appendChild(k)
                        },
                        updateState: function(e, g, f, k) {
                            h.item.call(k, f)
                        }
                    });
                    a.push({
                        type: _$_669e[517]
                    });
                    a.push({
                        content: _$_669e[23],
                        onclick: function(e, g, f) {
                            f.children[0].innerText == _$_669e[23] ? (r.current.fullscreen(!0), f.children[0].innerText = _$_669e[581]) : (r.current.fullscreen(!1), f.children[0].innerText = _$_669e[23])
                        },
                        tooltip: _$_669e[582],
                        updateState: function(e, g, f, k) {
                            f.children[0].innerText = 1 == k.parent.config.fullscreen ? _$_669e[581] : _$_669e[23]
                        }
                    });
                    a.push({
                        content: _$_669e[583],
                        onclick: function(e, g, f) {
                            r.current && r.current.parent.tools.search ? r.current.parent.tools.search.open(!0) : f.children[0].innerText == _$_669e[583] ? (xa.show.call(r.current),
                                f.children[0].innerText = _$_669e[584]) : (xa.hide.call(r.current), f.children[0].innerText = _$_669e[583])
                        },
                        tooltip: _$_669e[585],
                        updateState: function(e, g, f, k) {
                            f.children[0].innerText = 1 == k.options.search ? _$_669e[584] : _$_669e[583]
                        }
                    });
                    return a
                };
            return h
        }(),
        ua = function() {
            var h = function(c) {
                c.createWorksheet = function(b) {
                    return c.parent.createWorksheet(b)
                };
                c.deleteWorksheet = function(b) {
                    return c.parent.deleteWorksheet(b)
                };
                c.renameWorksheet = function(b, a) {
                    return c.parent.renameWorksheet(b, a)
                };
                c.moveWorksheet = function(b,
                    a) {
                    return c.parent.moveWorksheet(b, a)
                };
                c.openWorksheet = function() {
                    return c.parent.openWorksheet(c.parent.getWorksheet(c))
                };
                c.getWorksheet = function(b) {
                    return c.parent.getWorksheet(b)
                };
                c.getWorksheetActive = function() {
                    return c.parent.getWorksheetActive()
                };
                c.getWorksheetName = function() {
                    return c.options.worksheetName.toUpperCase()
                }
            };
            h.nextName = function() {
                var c = 1,
                    b, a = r.spreadsheet;
                if (a.length)
                    for (var d = 0; d < a.length; d++) {
                        var e = a[d].worksheets;
                        if (e.length)
                            for (var g = 0; g < e.length; g++) e[g].options.worksheetName &&
                                (b = e[g].options.worksheetName.match(/(\d+)/)) && (b = parseInt(b[0]) + 1, b > c && (c = b))
                    }
                return c
            };
            h.createWorksheet = function(c) {
                if (!T.call(this)) return !1;
                c || (c = {});
                var b = this.worksheets.length,
                    a = G.call(this, _$_669e[586], c, b);
                if (typeof a === _$_669e[109]) c = a;
                else if (!1 === a) return !1;
                c.data || c.minDimensions || (c.minDimensions = [8, 8]);
                c.worksheetId || (c.worksheetId = jSuites.guid().substring(0, 8));
                c.worksheetName || (c.worksheetName = _$_669e[422] + h.nextName());
                if (I(c.worksheetName)) jSuites.notification({
                    error: _$_669e[580],
                    message: _$_669e[587]
                });
                else return a = va.worksheet(this, c), a.onload(), r.current && r.current.resetSelection(), r.current = a, L.call(a, _$_669e[588], [c]), ma.update.call(this, r.current), G.call(this, _$_669e[589], a, c, b), a
            };
            h.deleteWorksheet = function(c) {
                var b = this.worksheets[c];
                if (!T.call(this, b)) return !1;
                I(b.options.worksheetName, null);
                this.element.tabs.deleteElement(c);
                this.worksheets.splice(c, 1);
                L.call(b, _$_669e[590], [c]);
                G.call(this, _$_669e[591], b, c)
            };
            h.renameWorksheet = function(c, b) {
                var a = this.worksheets[c];
                if (!T.call(this, a)) return !1;
                var d = a.options.worksheetName;
                if (d.toLowerCase() == b.toLowerCase()) return !1;
                b = b.replace(/[^0-9A-Za-z_\s^]+/gi, _$_669e[3]);
                Number(b) == b && (b = _$_669e[422] + Number(b));
                I(b) ? alert(_$_669e[592]) : (I(d, null), this.element.tabs.rename(c, b), a.options.worksheetName = b, I(b, a), Q.call(this, {
                    worksheet: a,
                    action: _$_669e[100],
                    index: c,
                    oldValue: d,
                    newValue: b
                }), L.call(a, _$_669e[100], {
                    worksheet: c,
                    newValue: b
                }), La.updateWorksheetName.call(a, d, b), Y.updateWorksheetNames.call(a, d, b), G.call(this, _$_669e[593],
                    a, c, b, d))
            };
            h.updateWorksheet = function(c, b) {
                var a = this.worksheets;
                if (!T.call(this)) return !1;
                this.worksheets.splice(b, 0, a.splice(c, 1)[0]);
                Q.call(this, {
                    worksheet: a[b],
                    action: _$_669e[101],
                    f: c,
                    t: b
                });
                L.call(a[b], _$_669e[101], {
                    f: c,
                    t: b
                });
                G.call(this, _$_669e[594], a[b], c, b)
            };
            h.moveWorksheet = function(c, b) {
                this.element.tabs.move(c, b)
            };
            h.openWorksheet = function(c) {
                var b = this.worksheets[c];
                !b || r.current && r.current == b || (this.element.tabs.getActive() != c ? this.element.tabs.open(c) : this.element.tabs.setBorder(c), r.current &&
                    r.current.resetSelection(), r.current = b, J.update.call(b), h.state.call(this, c), ma.update.call(this, r.current), G.call(this, _$_669e[595], b, c))
            };
            h.getWorksheet = function(c) {
                if (c && typeof c == _$_669e[109]) {
                    if (c = this.worksheets.indexOf(c), 0 <= c) return c
                } else
                    for (var b = 0; b < this.worksheets.length; b++)
                        if (c === this.worksheets[b].options.worksheetId) return b;
                return !1
            };
            h.getWorksheetActive = function() {
                return this.element.tabs.getActive()
            };
            h.getWorksheetInstance = function(c) {
                return c === _$_669e[3] || typeof c == _$_669e[25] ?
                    (c = h.getWorksheet.call(this, c), !1 === c ? !1 : this.worksheets[c]) : this.worksheets[c]
            };
            h.state = function(c) {
                try {
                    if (typeof localStorage !== _$_669e[21]) {
                        var b = this.config.cloud ? this.config.cloud : this.element.id;
                        if (b)
                            if (typeof c !== _$_669e[21]) localStorage.setItem(_$_669e[402] + b, c);
                            else {
                                var a = parseInt(localStorage.getItem(_$_669e[402] + b));
                                h.openWorksheet.call(this, a)
                            }
                    }
                } catch (d) {}
            };
            return h
        }(),
        H = function() {
            var h = {
                getCaretIndex: function(c) {
                    var b = 0,
                        a = (this.config.root ? this.config.root : window).getSelection();
                    a &&
                        0 !== a.rangeCount && (b = a.getRangeAt(0), a = b.cloneRange(), a.selectNodeContents(c), a.setEnd(b.endContainer, b.endOffset), b = a.toString().length);
                    return b
                },
                invert: function(c) {
                    for (var b = [], a = Object.keys(c), d = 0; d < a.length; d++) b[c[a[d]]] = a[d];
                    return b
                },
                getColumnName: function(c) {
                    var b = _$_669e[3];
                    701 < c ? (b += String.fromCharCode(64 + parseInt(c / 676)), b += String.fromCharCode(64 + parseInt(c % 676 / 26))) : 25 < c && (b += String.fromCharCode(64 + parseInt(c / 26)));
                    return b += String.fromCharCode(65 + c % 26)
                },
                getColumnNameFromCoords: function(c,
                    b) {
                    return h.getColumnName(parseInt(c)) + (parseInt(b) + 1)
                },
                getCoordsFromColumnName: function(c) {
                    if (!c) return [];
                    c = (_$_669e[3] + c).replace(new RegExp(/\$/), _$_669e[3]);
                    var b = /^[a-zA-Z]+/.exec(c);
                    if (b) {
                        for (var a = 0, d = 0; d < b[0].length; d++) a += parseInt(b[0].charCodeAt(d) - 64) * Math.pow(26, b[0].length - 1 - d);
                        a--;
                        0 > a && (a = 0);
                        c = parseInt(/[0-9]+$/.exec(c)) || null;
                        0 < c && c--;
                        return [a, c]
                    }
                }
            };
            h.shiftFormula = Y.shiftFormula;
            h.updateWorksheetName = Y.updateWorksheetName;
            h.createFromTable = function(c, b) {
                if (c.tagName != _$_669e[596]) console.log(_$_669e[597]);
                else {
                    b || (b = {});
                    b.columns = [];
                    b.data = [];
                    var a = c.querySelectorAll(_$_669e[598]);
                    if (a.length)
                        for (var d = 0; d < a.length; d++) {
                            var e = a[d].style.width;
                            e || (e = a[d].getAttribute(_$_669e[196]));
                            e && (b.columns[d] || (b.columns[d] = {}), b.columns[d].width = e)
                        }
                    e = function(B) {
                        var A = B.getBoundingClientRect();
                        A = 50 < A.width ? A.width : 50;
                        b.columns[d] || (b.columns[d] = {});
                        B.getAttribute(_$_669e[599]) ? b.columns[d].type = B.getAttribute(_$_669e[599]) : b.columns[d].type = _$_669e[69];
                        b.columns[d].width = A + _$_669e[57];
                        b.columns[d].title = B.innerHTML;
                        B.style.textAlign && (b.columns[d].align = B.style.textAlign);
                        B.classList.contains(_$_669e[600]) && (b.columns[d].readOnly = !0);
                        if (A = B.getAttribute(_$_669e[601])) b.columns[d].name = A;
                        if (A = B.getAttribute(_$_669e[602])) b.columns[d].id = A
                    };
                    var g = [],
                        f = c.querySelectorAll(_$_669e[603]);
                    if (f.length) {
                        for (a = 0; a < f.length - 1; a++) {
                            var k = [];
                            for (d = 0; d < f[a].children.length; d++) {
                                var l = {
                                    title: f[a].children[d].innerText,
                                    colspan: f[a].children[d].getAttribute(_$_669e[207]) || 1
                                };
                                k.push(l)
                            }
                            g.push(k)
                        }
                        f = f[f.length - 1].children;
                        for (d =
                            0; d < f.length; d++) e(f[d])
                    }
                    l = 0;
                    var n = {},
                        p = {},
                        q = {},
                        v = {},
                        w = null;
                    k = c.querySelectorAll(_$_669e[604]);
                    for (a = 0; a < k.length; a++)
                        if (b.data[l] = [], 1 != b.parseTableFirstRowAsHeader || f.length || 0 != a) {
                            for (d = 0; d < k[a].children.length; d++) {
                                var x = k[a].children[d].getAttribute(_$_669e[605]);
                                x ? x.substr(0, 1) != _$_669e[18] && (x = _$_669e[18] + x) : x = k[a].children[d].innerHTML;
                                b.data[l].push(x);
                                x = H.getColumnNameFromCoords(d, a);
                                (w = k[a].children[d].getAttribute(_$_669e[606])) && (v[x] = w);
                                w = parseInt(k[a].children[d].getAttribute(_$_669e[207])) ||
                                    0;
                                var y = parseInt(k[a].children[d].getAttribute(_$_669e[208])) || 0;
                                if (w || y) n[x] = [w || 1, y || 1];
                                if (w = k[a].children[d].style && k[a].children[d].style.display == _$_669e[52]) k[a].children[d].style.display = _$_669e[3];
                                (w = k[a].children[d].getAttribute(_$_669e[127])) && (q[x] = w);
                                k[a].children[d].classList.contains(_$_669e[607]) && (q[x] = q[x] ? q[x] + _$_669e[608] : _$_669e[609])
                            }
                            k[a].style && k[a].style.height && (p[a] = {
                                height: k[a].style.height
                            });
                            l++
                        } else
                            for (d = 0; d < k[a].children.length; d++) e(k[a].children[d]);
                    0 < Object.keys(g).length &&
                        (b.nestedHeaders = g);
                    0 < Object.keys(q).length && (b.style = q);
                    0 < Object.keys(n).length && (b.mergeCells = n);
                    0 < Object.keys(p).length && (b.rows = p);
                    0 < Object.keys(v).length && (b.classes = v);
                    k = c.querySelectorAll(_$_669e[610]);
                    if (k.length) {
                        c = [];
                        for (a = 0; a < k.length; a++) {
                            e = [];
                            for (d = 0; d < k[a].children.length; d++) e.push(k[a].children[d].innerText);
                            c.push(e)
                        }
                        0 < Object.keys(c).length && (b.footers = c)
                    }
                    if (1 == b.parseTableAutoCellType)
                        for (c = [], d = 0; d < b.columns.length; d++) {
                            g = e = !0;
                            c[d] = [];
                            for (a = 0; a < b.data.length; a++)
                                if (x = b.data[a][d],
                                    c[d][x] || (c[d][x] = 0), c[d][x]++, 25 < x.length && (e = !1), 10 == x.length) {
                                    if (x.substr(4, 1) != _$_669e[255] || x.substr(7, 1) != _$_669e[255]) g = !1
                                } else g = !1;
                            a = Object.keys(c[d]).length;
                            g ? b.columns[d].type = _$_669e[232] : 1 == e && 1 < a && a <= parseInt(.1 * b.data.length) && (b.columns[d].type = _$_669e[71], b.columns[d].source = Object.keys(c[d]))
                        }
                    return b
                }
            };
            h.injectArray = function(c, b, a) {
                return c.slice(0, b).concat(a).concat(c.slice(b))
            };
            h.parseCSV = function(c, b) {
                b = b || _$_669e[37];
                for (var a = 0, d = 0, e = [
                        []
                    ], g = 0, f = null, k = !1, l = !1, n = 0; n < c.length; n++)
                    if (e[d] ||
                        (e[d] = []), e[d][a] || (e[d][a] = _$_669e[3]), c[n] != _$_669e[611])
                        if (c[n] != _$_669e[612] && c[n] != b || 0 != k && 1 != l && f) {
                            c[n] == _$_669e[20] && (k = !k);
                            if (null === f) {
                                if (f = k, 1 == f) continue
                            } else if (!0 === f && !l && c[n] == _$_669e[20]) {
                                c[n + 1] == _$_669e[20] ? (k = !0, e[d][a] += c[n], n++) : l = !0;
                                continue
                            }
                            e[d][a] += c[n]
                        } else {
                            f = null;
                            l = k = !1;
                            if (e[d][a][0] == _$_669e[20]) {
                                var p = e[d][a].trim();
                                p[p.length - 1] == _$_669e[20] && (e[d][a] = p.substr(1, p.length - 2))
                            }
                            c[n] == _$_669e[612] ? (a = 0, d++) : (a++, a > g && (g = a))
                        } for (c = 0; c < e.length; c++)
                    for (n = 0; n <= g; n++) void 0 ===
                        e[c][n] && (e[c][n] = _$_669e[3]);
                return e
            };
            h.getTokensFromRange = Y.getTokensFromRange;
            h.getRangeFromTokens = Y.getRangeFromTokens;
            h.getCoordsFromRange = function(c) {
                -1 !== c.indexOf(_$_669e[19]) && (c = c.split(_$_669e[19]), c = c[1]);
                c = c.split(_$_669e[164]);
                c[1] || (c[1] = c[0]);
                var b = H.getCoordsFromColumnName(c[0]);
                c = H.getCoordsFromColumnName(c[1]);
                return [b[0], b[1], c[0], c[1]]
            };
            return h
        }(),
        Ga = function() {
            var h = {},
                c = null,
                b = null,
                a = null,
                d = null,
                e = !1,
                g = null,
                f = null,
                k = null,
                l = function(u) {
                    function z(S) {
                        S.className && (S.classList.contains(_$_669e[414]) &&
                            (K = S), S.classList.contains(_$_669e[425]) && (N = S), S.classList.contains(_$_669e[613]) && (V = S));
                        S.parentNode && S.parentNode.classList && S.parentNode.classList.contains(_$_669e[397]) && (S.tagName == _$_669e[614] ? C = 1 : S.tagName == _$_669e[615] ? C = 2 : S.tagName == _$_669e[616] && (C = 3));
                        S.parentNode && !K && z(S.parentNode)
                    }
                    var C = null,
                        E = null,
                        K = null,
                        N = null,
                        V = null,
                        O = null,
                        oa = null,
                        X = null;
                    z(u.target);
                    if (null !== V) return !1;
                    if (K)
                        if (u.target.classList.contains(_$_669e[395])) E = _$_669e[506];
                        else if (u.target.classList.contains(_$_669e[405])) E =
                        _$_669e[166];
                    else if (u.target.classList.contains(_$_669e[323])) E = _$_669e[469], O = parseInt(u.target.getAttribute(_$_669e[59]));
                    else if (u.target.parentNode && u.target.parentNode.classList && u.target.parentNode.classList.contains(_$_669e[224])) O = u.target.getAttribute(_$_669e[229]), oa = u.target.getAttribute(_$_669e[230]), null === O ? E = _$_669e[506] : (E = _$_669e[479], O = parseInt(O), oa = parseInt(oa));
                    else if (u.target.parentNode && u.target.parentNode.classList && u.target.parentNode.classList.contains(_$_669e[184])) E =
                        _$_669e[617];
                    else if (u.target.classList.contains(_$_669e[188])) E = _$_669e[617];
                    else if (u.target.parentNode && u.target.parentNode.classList && u.target.parentNode.classList.contains(_$_669e[618])) E = _$_669e[474], O = Array.prototype.indexOf.call(u.target.parentNode.children, u.target);
                    else if (1 == C) E = _$_669e[470], O = u.target.clientWidth - u.offsetX, u.target.classList.contains(_$_669e[215]) && 3 < O && 20 > O && (E = _$_669e[619]), O = u.target.getAttribute(_$_669e[63]);
                    else if (2 == C) {
                        if (u = n(u, K)) E = _$_669e[468], O = u[0], oa = u[1],
                            X = u[2]
                    } else if (3 == C) {
                        if (O = u.target.getAttribute(_$_669e[63]), oa = u.target.getAttribute(_$_669e[59]), O || oa) E = _$_669e[620]
                    } else C || jSuites.findElement(u.target, _$_669e[621]) && (E = _$_669e[112]);
                    return [K, N, C, E, O, oa, X]
                },
                n = function(u, z) {
                    if (u.changedTouches && u.changedTouches[0]) {
                        var C = u.changedTouches[0].clientX;
                        u = u.changedTouches[0].clientY
                    } else C = u.clientX, u = u.clientY;
                    z = z.spreadsheet.config.root ? z.spreadsheet.config.root : document;
                    if (z.elementsFromPoint) var E = z.elementsFromPoint(C, u);
                    else z.msElementsFromPoint &&
                        (E = z.msElementsFromPoint(C, u));
                    for (z = 0; z < E.length; z++)
                        if (C = E[z].getAttribute(_$_669e[63]), u = E[z].getAttribute(_$_669e[59]), null != C && null != u) return C = parseInt(C), u = parseInt(u), [C, u, E[z]];
                    return !1
                },
                p = function(u) {
                    var z = l(u);
                    if (!1 === z) return !1;
                    z[0] ? z[1] ? r.current && document.activeElement.classList.contains(_$_669e[121]) ? document.activeElement.worksheet || (document.activeElement.worksheet = z[1].jexcel) : r.current != z[1].jexcel && (r.current && r.current.resetSelection(), r.current = z[1].jexcel) : r.current || (u = z[0].tabs.getActive(),
                        0 <= u && (r.current = z[0].spreadsheet.worksheets[u])) : r.current && !jSuites.findElement(u.target, _$_669e[270]) && (r.current.resetSelection(!0), r.current = null);
                    return z
                };
            h.input = function(u) {
                if (u.target.tagName == _$_669e[622] && u.target.classList.contains(_$_669e[121])) {
                    var z = null;
                    u.target.worksheet ? z = u.target.worksheet : r.current && (z = r.current);
                    z && ta.parse.call(z, u.target)
                }
            };
            h.keydown = function(u) {
                if (r.current) {
                    u.target.tagName == _$_669e[622] && u.target.classList.contains(_$_669e[121]) && ta.onkeydown.call(r.current,
                        u);
                    if (27 == u.which) {
                        if (r.current.edition) {
                            r.current.closeEditor(r.current.edition, !1);
                            u.preventDefault();
                            return
                        }
                        if (ya.event) {
                            ya.cancel.call(r.current, u);
                            u.preventDefault();
                            return
                        }
                        if (fa.event) {
                            fa.cancel.call(r.current, u);
                            return
                        }
                        if (1 == la.isVisible()) {
                            la.close.call(r.current, !1);
                            u.preventDefault();
                            return
                        }
                    }
                    if (u.target.classList.contains(_$_669e[446])) {
                        var z = u.target;
                        g && clearTimeout(g);
                        g = setTimeout(function() {
                            la.search.call(r.current, z.value);
                            g = null
                        }, 500)
                    } else if (u.target.classList.contains(_$_669e[460])) z =
                        u.target, g && clearTimeout(g), g = setTimeout(function() {
                            r.current.search(z.value);
                            g = null
                        }, 400);
                    else if (!u.target.classList.contains(_$_669e[270]))
                        if (r.current.edition) {
                            var C = r.current.edition.getAttribute(_$_669e[63]);
                            13 == u.which ? r.current.options.columns[C].type == _$_669e[232] ? r.current.closeEditor(r.current.edition, !0) : r.current.options.columns[C].type != _$_669e[71] && r.current.options.columns[C].type != _$_669e[623] && r.current.options.columns[C].type != _$_669e[503] && (u.altKey ? r.current.parent.input.classList.contains(_$_669e[624]) ||
                                document.execCommand(_$_669e[625], !1, _$_669e[626]) : (r.current.closeEditor(r.current.edition, !0), r.current.parent.config.moveDownOnEnter && r.current.down(), u.preventDefault())) : 9 == u.which ? (r.current.closeEditor(r.current.edition, !0), r.current.right(), u.preventDefault()) : 39 != u.which && 37 != u.which && 38 != u.which && 40 != u.which || r.current.parent.input.selected || (r.current.closeEditor(r.current.edition, !0), 37 == u.which ? r.current.left(u.shiftKey, u.ctrlKey) : 39 == u.which ? r.current.right(u.shiftKey, u.ctrlKey) : 38 ==
                                u.which ? r.current.up(u.shiftKey, u.ctrlKey) : 40 == u.which && r.current.down(u.shiftKey, u.ctrlKey), u.preventDefault())
                        } else {
                            if ((u.ctrlKey || u.metaKey) && !u.shiftKey) {
                                if (65 == u.which) {
                                    r.current.selectAll();
                                    u.preventDefault();
                                    return
                                }
                                if (83 == u.which) {
                                    r.current.download();
                                    u.preventDefault();
                                    return
                                }
                                if (89 == u.which) {
                                    r.current.parent.redo();
                                    u.preventDefault();
                                    return
                                }
                                if (90 == u.which) {
                                    r.current.parent.undo();
                                    u.preventDefault();
                                    return
                                }
                                if (70 == u.which && r.current.parent.tools.search) {
                                    r.current.parent.tools.search.open(!0);
                                    u.preventDefault();
                                    return
                                }
                            }
                            if (r.current.selectedCell)
                                if (33 == u.which) 0 < r.current.options.pagination ? Ea.pageUp.call(r.current) : u.altKey ? J.pageLeft.call(r.current) : J.pageUp.call(r.current), u.preventDefault();
                                else if (34 == u.which) 0 < r.current.options.pagination ? Ea.pageDown.call(r.current) : u.altKey ? J.pageRight.call(r.current) : J.pageDown.call(r.current), u.preventDefault();
                            else if (37 == u.which) r.current.left(u.shiftKey, u.ctrlKey), u.preventDefault();
                            else if (39 == u.which) r.current.right(u.shiftKey, u.ctrlKey),
                                u.preventDefault();
                            else if (38 == u.which) r.current.up(u.shiftKey, u.ctrlKey), u.preventDefault();
                            else if (40 == u.which) r.current.down(u.shiftKey, u.ctrlKey), u.preventDefault();
                            else if (36 == u.which) r.current.first(u.shiftKey, u.ctrlKey), u.preventDefault();
                            else if (35 == u.which) r.current.last(u.shiftKey, u.ctrlKey), u.preventDefault();
                            else if (32 == u.which) r.current.setCheckRadioValue(), u.preventDefault();
                            else if (46 == u.which || 8 == u.which) r.current.isEditable() && r.current.setValue(r.current.getSelected(!1, !0), _$_669e[3]);
                            else if (13 == u.which) u.shiftKey ? r.current.up() : (1 == r.current.options.allowManualInsertRow && r.current.selectedCell[1] == (r.current.results ? r.current.results[r.current.results.length - 1] : r.current.rows.length - 1) && r.current.insertRow(null, null, null, null, u), r.current.selectedCell && r.current.down()), u.preventDefault();
                            else if (9 == u.which) u.shiftKey ? r.current.left() : (C = r.current.selectedCell[0], r.current.right(), 1 == r.current.options.allowInsertColumn && 1 == r.current.options.allowManualInsertColumn && r.current.selectedCell[0] ==
                                C && (r.current.insertColumn(null, null, null, null, null, null, u), r.current.right())), u.preventDefault();
                            else if ((u.ctrlKey || u.metaKey) && !u.shiftKey) 67 == u.which ? (r.current.copy(), u.preventDefault()) : 88 == u.which ? (r.current.cut(), u.preventDefault()) : 86 == u.which ? u.target.classList.contains(_$_669e[270]) || h.paste(u) : 189 == u.which && (1 == r.current.options.allowDeleteRow && k[3] == _$_669e[469] ? (ka.isValid.call(r.current, r.current.selectedCell) && r.current.deleteRow(), u.preventDefault()) : 1 == r.current.options.allowDeleteColumn &&
                                k[3] == _$_669e[470] && (ka.isValid.call(r.current, r.current.selectedCell) && r.current.deleteColumn(), u.preventDefault()));
                            else if (r.current.isEditable()) {
                                C = r.current.selectedCell[1];
                                var E = r.current.selectedCell[0];
                                r.current.options.columns[E].type != _$_669e[26] && (32 == u.keyCode ? r.current.options.columns[E].type == _$_669e[301] || r.current.options.columns[E].type == _$_669e[308] ? u.preventDefault() : r.current.openEditor(r.current.records[C][E].element, !0, u) : 113 == u.keyCode ? r.current.openEditor(r.current.records[C][E].element,
                                    !1, u) : 1 != u.key.length && u.key != _$_669e[627] || u.altKey || u.ctrlKey || r.current.records[C] && r.current.openEditor(r.current.records[C][E].element, !0, u))
                            }
                        }
                }
            };
            var q = [0, 0],
                v = function(u, z) {
                    z || (z = 800);
                    e && J.limited.call(r.current) && (f && (clearTimeout(f), f = null), f = setTimeout(function() {
                        if (e) {
                            200 < z && (z = 200);
                            z -= 10;
                            20 > z && (z = 20);
                            var C = q[0],
                                E = q[1],
                                K = r.current.content.getBoundingClientRect();
                            50 > K.height - (E - K.top) ? J.down.call(r.current) : 80 > E - K.top && J.up.call(r.current);
                            50 > K.width - (C - K.left) ? J.right.call(r.current) : 80 >
                                C - K.left && J.left.call(r.current);
                            v(u, z)
                        }
                        f = null
                    }, z))
                },
                w = function(u, z, C, E, K, N, V) {
                    if (document.activeElement.classList.contains(_$_669e[121]) && ta.update.call(r.current, u, z, C, E, K, document.activeElement, V)) return !1;
                    if (r.current.edition) r.current.closeEditor(r.current.edition, !0);
                    else if (d && d !== N) return !1;
                    r.current !== V && (r.current.resetSelection(), r.current = V);
                    ta.current ? ta.range.call(r.current, u, z, C, E, K) : ka[_$_669e[11]].call(r.current, u, z, C, E, K)
                };
            h.mousedown = function(u) {
                u = u || window.event;
                var z = u.buttons ?
                    u.buttons : u.button ? u.button : u.which;
                if (r.current)
                    if (1 == la.isVisible() && (jSuites.findElement(u.target, _$_669e[450]) || la.close.call(r.current, !1)), r.current.edition) {
                        if (jSuites.findElement(u.target, _$_669e[128])) return !1
                    } else if (jSuites.findElement(u.target, _$_669e[270])) return !1;
                var C = p(u);
                if (!1 === C) return !1;
                k = C;
                if (1 < z)
                    if (C[3] == _$_669e[468]) {
                        if (r.current.isSelected(C[4], C[5])) return !1
                    } else if (C[3] == _$_669e[469]) {
                    if (r.current.isSelected(0, C[4])) return !1
                } else if (C[3] == _$_669e[470] && r.current.isSelected(C[4],
                        0)) return !1;
                if (r.current) {
                    if (C[3] == _$_669e[166]) r.current.isEditable() && (fa.event = !0);
                    else if (C[3] == _$_669e[619]) la.open.call(r.current, C[4]);
                    else {
                        if (1 == C[2])
                            if (c) c.style.cursor == _$_669e[628] ? ya.start.call(r.current, u) : c.style.cursor == _$_669e[629] ? Da.start.call(r.current, u) : c.style.cursor == _$_669e[630] && (!0 === c.column ? ea.start.call(r.current, u, 0) : !1 === c.column ? ja.start.call(r.current, u, 0) : ea.start.call(r.current, u, parseInt(C[4]) + 1));
                            else {
                                z = r.current.options.columns.length - 1;
                                var E = r.current.options.data.length -
                                    1;
                                if (C[3] == _$_669e[506]) w(0, 0, z, E, u, C[3], C[1].jexcel), a = b = 0;
                                else if (C[3] == _$_669e[479]) {
                                    var K = Aa.range.call(r.current, C[5])[C[4]];
                                    K[1] > z && (K[1] = z);
                                    w(K[0], 0, K[1], E, u, C[3], C[1].jexcel);
                                    b = K[0];
                                    a = 0
                                } else C[4] && ((u.shiftKey || u.ctrlKey) && null != b && null != a ? w(b, 0, C[4], E, u, C[3], C[1].jexcel) : (C[3] == _$_669e[470] && b == C[4] && 1 == r.current.options.allowRenameColumn && (f = setTimeout(function() {
                                    r.current.setHeader(C[4])
                                }, 800)), w(C[4], 0, C[4], E, u, C[3], C[1].jexcel), b = C[4], a = 0))
                            } if (2 == C[2])
                            if (C[3] == _$_669e[469]) c ? c.style.cursor ==
                                _$_669e[628] ? ya.start.call(r.current, u) : c.style.cursor == _$_669e[631] ? Da.start.call(r.current, u) : c.style.cursor == _$_669e[630] && ja.start.call(r.current, u, parseInt(C[4]) + 1) : (z = r.current.options.columns.length - 1, E = parseInt(C[4]), (u.shiftKey || u.ctrlKey) && null != b && null != a ? w(0, a, z, E, u, C[3], C[1].jexcel) : (w(0, E, z, E, u, C[3], C[1].jexcel), b = 0, a = E));
                            else if (C[3] == _$_669e[468]) {
                            z = parseInt(C[4]);
                            E = parseInt(C[5]);
                            if (!r.current.edition && (K = r.current.options.columns[C[4]].type, u.target.tagName == _$_669e[632] && (K == _$_669e[303] ||
                                    K == _$_669e[304]) && u.target.parentNode.classList.contains(_$_669e[155]))) {
                                var N = u.target.innerText;
                                K == _$_669e[303] && (N = _$_669e[633] + N);
                                window.open(N, _$_669e[634])
                            }(u.shiftKey || u.ctrlKey) && r.current.selectedCell ? w(r.current.selectedCell[0], r.current.selectedCell[1], z, E, u, C[3], C[1].jexcel) : (w(z, E, z, E, u, C[3], C[1].jexcel), b = z, a = E)
                        }
                    }
                    d = C[3];
                    typeof C[3] !== _$_669e[21] && G.call(r.current.parent, _$_669e[635], r.current, C[3], C[4], C[5], u);
                    e = !0;
                    C[3] !== _$_669e[468] && C[3] !== _$_669e[166] || v(u)
                } else e = !1
            };
            h.mouseup =
                function(u) {
                    if (r.current)
                        if (ya.event) ya.end.call(r.current, u);
                        else if (Da.event) Da.end.call(r.current, u);
                    else if (ea.event) ea.end.call(r.current, u);
                    else if (ja.event) ja.end.call(r.current, u);
                    else if (fa.event) fa.end.call(r.current, u);
                    else if (u.target.classList.contains(_$_669e[449])) la.close.call(r.current, !0);
                    else if (r.current.edition && u.target.classList.contains(_$_669e[636]) && 50 > u.target.clientWidth - u.offsetX && 50 > u.offsetY) pa.close.call(r.current, null, !0);
                    else if (u.target.classList.contains(_$_669e[637])) {
                        let z =
                            u.target.offsetWidth - u.offsetX;
                        r.current.isEditable() && !r.current.edition && 8 < z && 18 > z && r.current.openEditor(u.target, null, u)
                    } else u.target.classList.contains(_$_669e[376]) && (u = u.target.innerText == _$_669e[15] ? 0 : u.target.innerText == _$_669e[17] ? u.target.getAttribute(_$_669e[315]) - 1 : u.target.innerText - 1, Ea[_$_669e[11]].call(r.current, parseInt(u)));
                    f && (clearTimeout(f), f = null);
                    e = d = !1;
                    ta.current && ta.close.call(r.current)
                };
            h.mousemove = function(u) {
                u = u || window.event;
                if (u.changedTouches && u.changedTouches[0]) {
                    var z =
                        u.changedTouches[0].clientX;
                    var C = u.changedTouches[0].clientY
                } else z = u.clientX, C = u.clientY;
                q = [z, C];
                if (r.current)
                    if (z = u.target.getAttribute(_$_669e[63]), C = u.target.getAttribute(_$_669e[59]), 1 == e) ya.event ? ya.update.call(r.current, u) : Da.event ? Da.update.call(r.current, u) : ea.event ? ea.move.call(r.current, u) : ja.event ? ja.move.call(r.current, u) : r.current.options.tableOverflow && (u = r.current.content, (u.w && u.w !== u.offsetWidth || u.h && u.h !== u.offsetHeight) && r.current.setViewport(u.offsetWidth, u.offsetHeight), u.w =
                        u.offsetWidth, u.h = u.offsetHeight);
                    else {
                        var E = u.target.getBoundingClientRect();
                        c && (c.style.cursor = _$_669e[3], c = null);
                        1 == r.current.options.editable && u.target && (u.target.classList.contains(_$_669e[395]) ? r.current.options.freezeRowControl && !r.current.options.freezeRows && 5 > E.height - (u.clientY - E.top) ? (c = u.target, c.style.cursor = _$_669e[630], c.column = !1) : r.current.options.freezeColumnControl && !r.current.options.freezeColumns && 5 > E.width - (u.clientX - E.left) && (c = u.target, c.style.cursor = _$_669e[630], c.column = !0) : u.target.classList.contains(_$_669e[323]) ? r.current.options.freezeRowControl && r.current.options.freezeRows == parseInt(C) + 1 && 5 > E.height - (u.clientY - E.top) ? (c = u.target, c.style.cursor = _$_669e[630]) : 1 == r.current.options.rowDrag && 6 > E.width - (u.clientX - E.left) ? (c = u.target, c.style.cursor = _$_669e[628]) : 1 == r.current.options.rowResize && 6 > E.height - (u.clientY - E.top) && (c = u.target, c.style.cursor = _$_669e[631]) : u.target.classList.contains(_$_669e[212]) && (r.current.options.freezeColumnControl && r.current.options.freezeColumns ==
                            parseInt(z) + 1 && 5 > E.width - (u.clientX - E.left) ? (c = u.target, c.style.cursor = _$_669e[630]) : 1 == r.current.options.columnResize && 6 > E.width - (u.clientX - E.left) ? (c = u.target, c.style.cursor = _$_669e[629]) : 1 == r.current.options.columnDrag && 6 > E.height - (u.clientY - E.top) && (c = u.target, c.style.cursor = _$_669e[628])))
                    }
            };
            h.mouseover = function(u) {
                u = u || window.event;
                if (r.current && 1 == e) {
                    var z = l(u);
                    if (z[0]) {
                        if (!document.activeElement.classList.contains(_$_669e[121]) && z[1] && r.current != z[1].jexcel && r.current) return !1;
                        if (!(ya.event ||
                                Da.event || ea.event || ja.event))
                            if (fa.event) {
                                var C = u.target.getAttribute(_$_669e[63]),
                                    E = u.target.getAttribute(_$_669e[59]);
                                r.current.isSelected(C, E) ? fa.reset.call(r.current) : null !== C && null !== E && fa.call(r.current, C, E)
                            } else null !== b && null !== a && (1 == z[2] && (C = b, z[3] == _$_669e[470] && (C = u.target.getAttribute(_$_669e[63])), E = r.current.options.data.length - 1, w(b, 0, C, E, u, z[3], z[1].jexcel), f && (clearTimeout(f), f = null)), 2 == z[2] && (z[3] == _$_669e[469] ? (C = r.current.options.columns.length - 1, E = parseInt(u.target.getAttribute(_$_669e[59])),
                                w(0, a, C, E, u, z[3], z[1].jexcel)) : z[3] == _$_669e[468] && (C = z[4], E = z[5], w(b, a, C, E, u, z[3], z[1].jexcel))))
                    }
                }
            };
            h.dblclick = function(u) {
                if (r.current)
                    if (u.target.classList.contains(_$_669e[405])) fa.execute.call(r.current, !1);
                    else {
                        var z = l(u);
                        1 == z[2] ? (u = u.target.getAttribute(_$_669e[63])) && (c && c.style.cursor == _$_669e[629] ? r.current.autoWidth(r.getSelectedColumns()) : 1 == r.current.options.columnSorting && 1 == r.current.options.columnSortingOnDblClick && r.current.orderBy(u)) : 2 == z[2] && r.current.isEditable() && !r.current.edition &&
                            z[6] && r.current.openEditor(z[6], null, u)
                    }
            };
            h.paste = function(u) {
                if (!u.target.classList.contains(_$_669e[270]) && r.current && r.current.selectedCell && !r.current.edition) {
                    var z = u.clipboardData || window.clipboardData;
                    z && (r.current.paste(Math.min(r.current.selectedCell[0], r.current.selectedCell[2]), Math.min(r.current.selectedCell[1], r.current.selectedCell[3]), z.getData(_$_669e[69])), u && u.preventDefault())
                }
            };
            h.contextmenu = function(u) {
                var z = p(u);
                !1 !== z && (z[1] || (z = k), r.current && !r.current.edition && (z = Array.prototype.slice.call(z,
                    3), Xa.open.call(r.current, u, z)))
            };
            var x = null,
                y = null,
                B = null,
                A = null,
                F = null;
            h.touchstart = function(u) {
                var z = l(u);
                z[0] ? z[1] && r.current != z[1].jexcel && (r.current && r.current.resetSelection(), r.current = z[1].jexcel) : r.current && (r.current.resetSelection(), r.current = null);
                r.current && !r.current.edition ? z[3] == _$_669e[166] ? r.current.isEditable() && (fa.event = !0) : u.target && u.target.classList.contains(_$_669e[169]) ? F = parseInt(u.target.getAttribute(_$_669e[170])) : z[3] == _$_669e[468] ? (A = z, x === u.target && y ? r.current.openEditor(u.target,
                    !1, u) : (y = setTimeout(function() {
                    x = y = A = null
                }, 400), B = setTimeout(function() {
                    h.contextmenu(u);
                    B = null
                }, 800)), x = u.target) : A = null : A = null
            };
            h.touchmove = function(u) {
                if (r.current) {
                    var z = null,
                        C = null,
                        E = null,
                        K = null;
                    if (u.changedTouches && u.changedTouches[0]) {
                        var N = u.changedTouches[0].clientX;
                        var V = u.changedTouches[0].clientY
                    } else N = u.clientX, V = u.clientY;
                    if (V = document.elementFromPoint(N, V))
                        if (N = V.getAttribute(_$_669e[63]), V = V.getAttribute(_$_669e[59]), null !== N && null !== V && fa.event) r.current.isSelected(N, V) || (fa.call(r.current,
                            N, V), v(u)), u.preventDefault();
                        else if (null !== F) {
                        N = parseInt(N);
                        V = parseInt(V);
                        var O = r.current.selectedCell;
                        0 == F ? N <= O[2] && (z = N, E = O[2], C = O[1], K = O[3]) : 2 == F ? N >= O[0] && (z = O[0], C = O[1], E = N, K = O[3]) : 1 == F ? V <= O[3] && (z = O[0], C = V, E = O[2], K = O[3]) : 3 == F && V >= O[1] && (z = O[0], C = O[1], E = O[2], K = V);
                        null !== z && r.current.updateSelectionFromCoords(z, C, E, K);
                        u.preventDefault()
                    } else A = null;
                    B && clearInterval(B)
                }
            };
            h.touchend = function(u) {
                if (fa.event) fa.end.call(r.current, u);
                else if (null !== F) F = null;
                else if (A) {
                    u = null;
                    if (A[3] === _$_669e[468]) {
                        var z =
                            A[4];
                        var C = A[5];
                        var E = z;
                        u = C
                    } else A[3] === _$_669e[469] ? (z = 0, C = A[4], E = r.current.options.columns.length - 1, u = C) : A[3] === _$_669e[470] && (z = A[4], C = 0, E = z, u = r.current.rows.length - 1);
                    null != z && r.current.updateSelectionFromCoords(z, C, E, u)
                }
                B && clearInterval(B)
            };
            h.resize = function() {
                r.current && 1 == r.current.parent.config.fullscreen && J.call(r.current, !0)
            };
            return h
        }(),
        r = function(h, c, b) {
            if (h) {
                if (h.tagName == _$_669e[596]) {
                    c = r.helpers.createFromTable(h, c);
                    var a = document.createElement(_$_669e[32]);
                    h.parentNode.insertBefore(a,
                        h);
                    h.parentNode.removeChild(h);
                    h = a
                }
                c || (c = {});
                r.defaultRowHeight || (a = document.createElement(_$_669e[32]), a.innerHTML = _$_669e[639], document.body.appendChild(a), r.defaultRowHeight = a.querySelector(_$_669e[185]).offsetHeight, a.parentNode.removeChild(a));
                var d = [];
                if (c.cloud && !b) return typeof r.cloud == _$_669e[110] ? r.cloud.load(c, function(g) {
                    if ((g = r(h, g, !0)) && g.length)
                        for (var f = 0; f < g.length; f++) d.push(g[f])
                }) : console.error(_$_669e[640]), d;
                if (h.spreadsheet) e = h.spreadsheet, Ha[_$_669e[11]].call(e, c);
                else var e =
                    va.spreadsheet(h, c);
                e.processing = !0;
                1 == e.config.loadingSpin && jSuites.loading.show();
                jSuites.ajax.oncomplete[e.name] = function() {
                    for (var g = 0; g < e.worksheets.length; g++) typeof e.worksheets[g].onload == _$_669e[110] && (e.worksheets[g].onload(), e.worksheets[g].onload = null);
                    e.processing = !1;
                    if (Array.isArray(e.queue) && 0 < e.queue.length)
                        for (; g = e.queue.shift();) g[0].updateCell(g[1], g[2], R.call(g[0], g[1], g[2]), !0);
                    e.config.toolbar && ma.show.call(e);
                    for (g = 0; g < e.worksheets.length; g++) ia.refresh.call(e.worksheets[g]);
                    G.call(e, _$_669e[641], e);
                    for (g = 0; g < e.worksheets.length; g++) ha.update.call(e.worksheets[g]);
                    jSuites.loading.hide()
                };
                Array.isArray(c.worksheets) && (c = c.worksheets);
                if (Array.isArray(c)) {
                    for (b = 0; b < c.length; b++) va.worksheet(e, c[b]);
                    d = e.worksheets
                } else d = va.worksheet(e, c);
                jSuites.ajax.pending(e.name) || (jSuites.ajax.oncomplete[e.name](), jSuites.ajax.oncomplete[e.name] = null);
                return d
            }
            console.error(_$_669e[638])
        };
    r.defaultRowHeight = null;
    r.license = null;
    r.setLicense = function(h) {
        r.license = h
    };
    r.spreadsheet = [];
    r.picker = ta;
    r.setDictionary = function(h) {
        jSuites.setDictionary(h)
    };
    r.extensions = {};
    r.setExtensions = function(h) {
        for (var c = Object.keys(h), b = 0; b < c.length; b++) typeof h[c[b]] === _$_669e[110] && (r[c[b]] = h[c[b]], r.extensions[c[b]] = h[c[b]], r.license && typeof r[c[b]].license == _$_669e[110] && r[c[b]].license.call(r, r.license))
    };
    r.destroy = function(h, c) {
        if (h.spreadsheet) {
            c && va.unbind(h);
            c = Object.keys(r.extensions);
            if (c.length)
                for (var b = 0; b < c.length; b++)
                    if (typeof r.extensions[c[b]].ondestroy == _$_669e[110]) r.extensions[c[b]].ondestroy(h.spreadsheet);
            c = r.spreadsheet.indexOf(h.spreadsheet);
            r.spreadsheet.splice(c, 1);
            for (b = 0; b < h.spreadsheet.worksheets.length; b++) I(h.spreadsheet.worksheets[b].options.worksheetName, null);
            h.spreadsheet = null;
            h.jexcel = null;
            h.innerHTML = _$_669e[3];
            h.removeAttribute(_$_669e[606])
        }
    };
    r.version = M;
    r.helpers = H;
    r.events = Ga;
    typeof formula !== _$_669e[21] && (r.formula = formula);
    r.editors = function() {
        var h = {},
            c = function(b, a) {
                return 1 == b[a] || b.options && 1 == b.options[a] ? !0 : b[a] == _$_669e[279] || b.options && b.options[a] == _$_669e[279] ? !0 : !1
            };
        h.createEditor = function(b, a, d, e) {
            e.parent.input.setAttribute(_$_669e[129], !1);
            b == _$_669e[269] ? (b = document.createElement(_$_669e[269]), b.type = _$_669e[69], b.value = d) : b = document.createElement(_$_669e[32]);
            b.style.width = a.offsetWidth - 2 + _$_669e[57];
            b.style.height = a.offsetHeight - 1 + _$_669e[57];
            e.parent.input.appendChild(b);
            return b
        };
        h.createDialog = function(b, a, d, e, g) {
            b = h.createEditor(_$_669e[32], b, a, g);
            b.classList.add(_$_669e[642]);
            b.classList.add(_$_669e[636]);
            800 > window.innerWidth ? (g.parent.input.style.top =
                _$_669e[173], g.parent.input.style.left = _$_669e[173]) : (a = g.parent.input.getBoundingClientRect(), window.innerHeight < a.bottom && (g.parent.input.style.marginTop = -g.parent.input.offsetHeight));
            return b
        };
        h.general = function() {
            var b = {},
                a = function(e) {
                    return (_$_669e[3] + e)[0] == _$_669e[18] ? !0 : !1
                },
                d = function(e) {
                    if (e.format || e.mask || e.locale) {
                        var g = {};
                        e.mask ? g.mask = e.mask : e.format ? g.mask = e.format : g.locale = e.locale;
                        e.options && (g.options = e.options);
                        e.decimal && (g.options || (g.options = {}), g.options = {
                            decimal: e.decimal
                        });
                        return g.mask === _$_669e[643] ? !1 : g
                    }
                    return null
                };
            b.createCell = function(e, g, f, k, l, n) {
                return b.updateCell(e, g, f, k, l, n)
            };
            b.updateCell = function(e, g, f, k, l, n) {
                f = parseInt(f);
                k = parseInt(k);
                if (e) {
                    var p = d(n),
                        q = null;
                    g === _$_669e[3] || a(g) || typeof g === _$_669e[22] || typeof g != _$_669e[25] || g.trim() === _$_669e[3] || (p ? (p = jSuites.mask.extract(g, p, !0)) && p.value !== _$_669e[3] && (q = p.value) : 1 == (!1 === l.parent.config.autoCasting || !1 === n.autoCasting ? !1 : !0) && (n.type == _$_669e[22] || n.type == _$_669e[365] || n.type == _$_669e[364] ? (p = jSuites.mask.extract(g,
                        n, !0)) && p.value !== _$_669e[3] && (q = p.value) : jSuites.isNumeric(Number(g)) && (q = Number(g))));
                    null !== q && (g = q);
                    g = b.parseValue(f, k, g, l, n, e);
                    g instanceof Element || g instanceof HTMLDocument ? (e.innerHTML = _$_669e[3], e.appendChild(g)) : n.rotate ? (e.firstChild && e.firstChild.classList.contains(_$_669e[27]) ? l = e.firstChild : (l = document.createElement(_$_669e[28]), l.classList.add(_$_669e[27])), l.style.transform = _$_669e[29] + parseInt(270 - n.rotate) + _$_669e[30], l.innerHTML = g, e.appendChild(l)) : 1 == l.parent.config.stripHTML ?
                        e.innerText = g : e.innerHTML = g;
                    if (null !== q) return q
                }
            };
            b.openEditor = function(e, g, f, k, l, n) {
                var p = l.parent.input;
                p.onblur = function() {
                    p.classList.contains(_$_669e[121]) || l.closeEditor(e, !0)
                };
                if (n.type == _$_669e[22] || n.type == _$_669e[365] || n.type == _$_669e[364]) p.classList.add(_$_669e[624]), a(g) || n.inputmode || (n.inputmode = _$_669e[644]);
                f = null;
                a(g) || (n.inputmode && p.setAttribute(_$_669e[125], n.inputmode), (f = d(n)) ? n.disabledMaskOnEdition || (n.mask ? (k = n.mask.split(_$_669e[262]), p.setAttribute(_$_669e[123], k[0])) :
                    n.locale && p.setAttribute(_$_669e[645], n.locale), f.input = p, p.mask = f, jSuites.mask.render(g, f, !1)) : g && n.type == _$_669e[364] && (g *= 100));
                if (!f || n.disabledMaskOnEdition) l.parent.config.stripHTML ? p.innerText = g : p.innerHTML = g, jSuites.focus(p)
            };
            b.closeEditor = function(e, g, f, k, l, n) {
                e = l.parent.input;
                g = g ? e.innerText : _$_669e[3];
                a(g) ? g = g.replace(/(\r\n|\n|\r)/gm, _$_669e[3]) : n.type == _$_669e[364] && (g += _$_669e[646], e.classList.remove(_$_669e[647]));
                return g
            };
            b.parseValue = function(e, g, f, k, l, n) {
                1 == k.parent.config.parseFormulas &&
                    (k.resetArray(e, g), a(f) && (f = k.executeFormula(f, e, g)));
                f instanceof Element || f instanceof HTMLDocument || (n && (!l.rotate && jSuites.isNumeric(f) ? n.classList.add(_$_669e[297]) : n.classList.remove(_$_669e[297])), l && !a(f) && (e = d(l)) && (e.mask ? (g = e.mask.split(_$_669e[262]), 0 > f && g[1] ? (n && -1 !== g[1].indexOf(_$_669e[648]) && n.classList.add(_$_669e[649]), n = -1 === g[1].indexOf(_$_669e[255]) ? jSuites.mask.render(-1 * f, e, !0) : jSuites.mask.render(f, e, !0)) : (e.mask = g[0], n && n.classList.remove(_$_669e[649]), n = jSuites.mask.render(f,
                    e, !0))) : n = jSuites.mask.render(f, e, !0), n && (f = n)));
                return f
            };
            b[_$_669e[56]] = function(e, g) {
                var f;
                return (f = d(e)) ? jSuites.mask.render(g, f, !0) : e.type == _$_669e[364] ? 100 * parseFloat(g) : g
            };
            b.destroyCell = function(e) {
                e.classList.remove(_$_669e[297])
            };
            return b
        }();
        h.text = h.general;
        h.number = h.general;
        h.numeric = h.general;
        h.percent = h.general;
        h.notes = function() {
            var b = {},
                a = null;
            b.updateCell = function(d, e, g, f, k, l) {
                d && (g = document.createElement(_$_669e[32]), g.classList.add(_$_669e[650]), 1 == k.parent.config.stripHTML ? g.innerText =
                    e : g.innerHTML = e, d.innerHTML = _$_669e[3], d.appendChild(g))
            };
            b.createCell = b.updateCell;
            b.openEditor = function(d, e, g, f, k, l) {
                var n = l && l.options ? l.options : {};
                n.focus = !0;
                n.value = e;
                n.border = !1;
                n.height = _$_669e[651];
                n.toolbar = !1;
                a = h.createDialog(d, e, g, f, k, l);
                a = jSuites.editor(a, n)
            };
            b.closeEditor = function(d, e, g, f, k, l) {
                return e ? a.getData() : _$_669e[3]
            };
            return b
        }();
        h.dropdown = function() {
            var b = {
                createCell: function(a, d, e, g, f, k) {
                    k.render == _$_669e[652] ? a.classList.add(_$_669e[653]) : a.classList.add(_$_669e[637]);
                    if (a =
                        b.updateCell(a, d, e, g, f, k)) return a
                },
                destroyCell: function(a) {
                    a.classList.remove(_$_669e[637])
                },
                updateCell: function(a, d, e, g, f, k) {
                    if (a) {
                        var l = b.getItem(a, d, e, g, f, k);
                        l ? a.innerHTML = l : d ? (d = b.fromLabel(a, d, e, g, f, k), d.length ? (l = b.getItem(a, d, e, g, f, k), a.innerHTML = l) : a.innerHTML = _$_669e[3]) : a.innerHTML = _$_669e[3];
                        return d
                    }
                },
                openEditor: function(a, d, e, g, f, k) {
                    var l = h.createEditor(_$_669e[32], a, d, f),
                        n = k.delimiter || _$_669e[262];
                    k = Object.create(k);
                    k.options || (k.options = {});
                    typeof k.filter == _$_669e[110] && (k.source =
                        k.filter(f.el, a, e, g, k.source));
                    typeof k.filterOptions == _$_669e[110] && (k = k.filterOptions(f, a, e, g, d, k));
                    e = k.options;
                    k.source && (e.data = JSON.parse(JSON.stringify(k.source)));
                    c(k, _$_669e[654]) && (e.multiple = !0);
                    c(k, _$_669e[70]) && (e.autocomplete = !0);
                    e.format = !0;
                    e.opened = !0;
                    e.width = a.offsetWidth - 2;
                    e.onclose = function() {
                        f.closeEditor(a, !0)
                    };
                    c(k, _$_669e[654]) ? d && (e.value = (_$_669e[3] + d).split(n)) : e.value = d;
                    jSuites.dropdown(l, e)
                },
                closeEditor: function(a, d, e, g, f, k) {
                    for (var l = f.parent.input.children[0], n = k.delimiter ||
                            _$_669e[262], p = l.dropdown.getText(!0), q = l.dropdown.getValue(!0), v = 0; v < q.length; v++) !1 === b.getItem(a, q[v], e, g, f, k) && k.source.push({
                        id: q[v],
                        name: p[v]
                    });
                    l.dropdown.close(!0);
                    if (d) return q.join(n)
                },
                fromLabel: function(a, d, e, g, f, k) {
                    a = k.delimiter || _$_669e[262];
                    e = [];
                    g = [];
                    Array.isArray(d) || (d = (_$_669e[3] + d).split(a));
                    for (f = 0; f < d.length; f++) g[d[f].trim()] = !0;
                    d = Object.keys(g);
                    !c(k, _$_669e[654]) && 1 < Object.keys(d).length && (g = [], g[d[0]] = !0);
                    for (d = 0; d < k.source.length; d++) g[k.source[d].name] && e.push(k.source[d].id);
                    return e.join(a)
                },
                getItem: function(a, d, e, g, f, k) {
                    g = k.delimiter || _$_669e[262];
                    a = [];
                    e = [];
                    Array.isArray(d) || (d = (_$_669e[3] + d).split(g));
                    for (g = 0; g < d.length; g++) e[d[g].trim()] = !0;
                    g = Object.keys(e);
                    !c(k, _$_669e[654]) && 1 < Object.keys(g).length && (d = [], d[g[0]] = !0, e = d);
                    if (k.source && typeof k.source[0] !== _$_669e[109])
                        for (d = 0; d < k.source.length; d++) g = k.source[d], k.source[d] = {
                            id: g,
                            name: g
                        };
                    for (d = 0; d < k.source.length; d++) typeof e[k.source[d].id] !== _$_669e[21] && a.push(k.render == _$_669e[305] ? k.source[d].color : k.render ==
                        _$_669e[655] ? k.source[d].image : k.source[d].name);
                    if (a.length) {
                        if (typeof k.render !== _$_669e[21]) {
                            for (d = 0; d < a.length; d++) e = d, g = a[d], g = k.render == _$_669e[305] ? _$_669e[656] + g + _$_669e[657] : k.render == _$_669e[655] ? _$_669e[656] + g + _$_669e[657] : k.render == _$_669e[652] ? _$_669e[659] + (k.source[d].color || _$_669e[658]) + _$_669e[660] + g + _$_669e[661] : void 0, a[e] = g;
                            return a.join(_$_669e[3])
                        }
                        for (d = 0; d < a.length; d++) a[d] = (_$_669e[3] + a[d]).replace(_$_669e[15], _$_669e[13]);
                        return a.join(_$_669e[353])
                    }
                    return !1
                }
            };
            b[_$_669e[56]] =
                function(a, d) {
                    for (var e = 0; e < a.source.length; e++)
                        if (a.source[e].id == d) return a.render == _$_669e[305] ? a.source[e].color : a.render == _$_669e[655] ? a.source[e].image : a.source[e].name
                };
            return b
        }();
        h.autocomplete = h.dropdown;
        h.calendar = function() {
            var b = {},
                a = null;
            b.updateCell = function(d, e, g, f, k, l) {
                if (d) {
                    g = e;
                    l = b.getFormat(l);
                    0 < e && Number(e) == e && (g = jSuites.calendar.numToDate(e));
                    f = _$_669e[3] + e;
                    f.substr(4, 1) == _$_669e[255] && f.substr(7, 1) == _$_669e[255] ? f = !0 : (f = f.split(_$_669e[255]), f = 4 == f[0].length && f[0] == Number(f[0]) &&
                        2 == f[1].length && f[1] == Number(f[1]) ? !0 : !1);
                    f || (f = jSuites.calendar.extractDateFromString(e, l)) && (g = f);
                    if (g = jSuites.calendar.getDateString(g, l)) return d.innerText = g, e;
                    d.innerText = _$_669e[3];
                    return _$_669e[3]
                }
            };
            b.createCell = b.updateCell;
            b.openEditor = function(d, e, g, f, k, l) {
                l = Object.create(l);
                typeof l.filterOptions == _$_669e[110] && (l = l.filterOptions(k, d, g, f, e, l));
                g = l.options || {};
                g.opened = !0;
                g.onclose = function(n, p) {
                    k.closeEditor(d, !0)
                };
                typeof l.timepicker !== _$_669e[21] && (g.time = c(l, _$_669e[662]) ? !0 : !1);
                typeof g.readonly !==
                    _$_669e[21] && (g.readonly = g.readonly);
                g.value = e || null;
                g.placeholder = g.format;
                a = h.createEditor(_$_669e[269], d, e, k);
                jSuites.calendar(a, g);
                !1 === a.calendar.options.readonly && a.focus()
            };
            b.closeEditor = function(d, e, g, f, k) {
                d = a.calendar.close(!0, a.value ? !0 : !1);
                d = a.value ? d : _$_669e[3];
                if (e) return d
            };
            b.getFormat = function(d) {
                return d && d.format ? d.format : d && d.options && d.options.format ? d.options.format : _$_669e[663]
            };
            b[_$_669e[56]] = function(d, e) {
                d = b.getFormat(d);
                return jSuites.calendar.getDateString(e, d)
            };
            return b
        }();
        h.color = function() {
            var b = {},
                a = null;
            b.updateCell = function(d, e, g, f, k, l) {
                d && (l.render == _$_669e[664] ? (g = document.createElement(_$_669e[32]), g.className = _$_669e[305], g.style.backgroundColor = e, d.innerText = _$_669e[3], d.appendChild(g)) : (d.style.color = e, d.innerText = e))
            };
            b.createCell = b.updateCell;
            b.openEditor = function(d, e, g, f, k, l) {
                l = Object.create(l);
                typeof l.filterOptions == _$_669e[110] && (l = l.filterOptions(k, d, g, f, e, l));
                g = l.options || {};
                g.value = e;
                g.onclose = function(n, p) {
                    k.closeEditor(d, !0)
                };
                a = h.createEditor(_$_669e[269],
                    d, e, k);
                a = jSuites.color(a, g);
                a.open()
            };
            b.closeEditor = function(d, e, g, f, k, l) {
                d = a.close(!0);
                if (e) return d
            };
            b[_$_669e[56]] = function(d, e) {
                d = document.createElement(_$_669e[32]);
                d.style.width = _$_669e[665];
                d.style.height = _$_669e[34];
                d.style.backgroundColor = e;
                return d.outerHTML
            };
            return b
        }();
        h.checkbox = function() {
            var b = {
                createCell: function(a, d, e, g, f, k) {
                    d = d && d != _$_669e[666] && d != _$_669e[667] ? !0 : !1;
                    e = document.createElement(_$_669e[269]);
                    e.type = _$_669e[301];
                    e.checked = d;
                    e.onclick = function(l) {
                        l.preventDefault();
                        l.stopPropagation();
                        if (k && 1 == k.readOnly || !f.isEditable()) return !1;
                        var n = this.checked;
                        setTimeout(function() {
                            f.setValue(a, n)
                        })
                    };
                    a.innerHTML = _$_669e[3];
                    a.appendChild(e);
                    return d
                },
                updateCell: function(a, d) {
                    d = d && d != _$_669e[666] && d != _$_669e[667] ? !0 : !1;
                    a && a.children[0] && (a.children[0].checked = d);
                    return d
                },
                openEditor: function(a, d, e, g, f) {
                    d = a.children[0].checked ? !1 : !0;
                    f.setValue(a, d);
                    return !1
                },
                closeEditor: function(a, d) {
                    return !1
                }
            };
            b[_$_669e[56]] = function(a, d) {
                return 1 == d || 1 == d || d == _$_669e[279] ? jSuites.translate(_$_669e[279]) : jSuites.translate(_$_669e[666])
            };
            return b
        }();
        h.radio = function() {
            var b = {
                createCell: function(a, d, e, g, f, k) {
                    a.getAttribute(_$_669e[63]);
                    e = document.createElement(_$_669e[269]);
                    e.type = _$_669e[308];
                    e.checked = 1 == d || 1 == d || d == _$_669e[279] ? !0 : !1;
                    e.onclick = function(l) {
                        l.preventDefault();
                        l.stopPropagation();
                        if (k && 1 == k.readOnly || !f.isEditable()) return !1;
                        setTimeout(function() {
                            f.setValue(a, !0)
                        })
                    };
                    a.innerHTML = _$_669e[3];
                    a.appendChild(e)
                },
                updateCell: function(a, d, e, g, f) {
                    d = 1 == d || 1 == d || d == _$_669e[279] || d == _$_669e[668] ? !0 : !1;
                    a && (a.children[0].checked =
                        d);
                    if (1 == d) {
                        a = [];
                        for (var k = 0; k < f.options.data.length; k++) g != k && f.options.data[k][e] && a.push({
                            x: e,
                            y: k,
                            value: 0
                        });
                        a.length && f.setValue(a)
                    }
                    return d
                },
                openEditor: function(a, d, e, g, f) {
                    return !1
                },
                closeEditor: function(a, d) {
                    return !1
                }
            };
            b[_$_669e[56]] = function(a, d) {
                return 1 == d || 1 == d || d == _$_669e[279] ? jSuites.translate(_$_669e[279]) : jSuites.translate(_$_669e[666])
            };
            return b
        }();
        h.autonumber = function() {
            var b = {},
                a = null;
            b.createCell = function(d, e, g, f, k, l) {
                (e = parseInt(e)) || d && 0 < parseInt(d.innerText) && (e = parseInt(d.innerText));
                l.sequence || (l.sequence = 0);
                e ? b.isValid(e, g, f, k) || (e = _$_669e[293]) : e = l.sequence + 1;
                e > l.sequence && (l.sequence = e);
                d && (d.innerText = e);
                return e
            };
            b.updateCell = b.createCell;
            b.openEditor = function(d, e, g, f, k, l) {
                a = h.createEditor(_$_669e[269], d, e, k);
                a.onblur = function() {
                    k.closeEditor(d, !0)
                };
                a.focus();
                a.value = e
            };
            b.closeEditor = function(d, e, g, f, k) {
                return e ? a.value : _$_669e[3]
            };
            b.isValid = function(d, e, g, f) {
                for (var k, l = 0; l < f.options.data.length; l++)
                    if (k = f.value(e, l), k == d && l != g) return !1;
                return !0
            };
            return b
        }();
        h.progressbar =
            function() {
                var b = {},
                    a = null;
                b.createCell = function(d, e) {
                    e = parseInt(e);
                    if (100 < e) e = 100;
                    else if (!e || 0 > e) e = 0;
                    if (d.children[0] && d.children[0].tagName == _$_669e[622]) var g = d.children[0];
                    else g = document.createElement(_$_669e[32]), g.classList.add(_$_669e[306]), d.innerText = _$_669e[3], d.classList.add(_$_669e[669]), d.appendChild(g);
                    g.style.width = parseInt(e) + _$_669e[646];
                    g.setAttribute(_$_669e[315], parseInt(e) + _$_669e[646])
                };
                b.destroyCell = function(d) {
                    d.classList.remove(_$_669e[669])
                };
                b.updateCell = b.createCell;
                b.openEditor =
                    function(d, e, g, f, k, l) {
                        a = h.createEditor(_$_669e[269], d, e, k);
                        a.type = _$_669e[22];
                        a.setAttribute(_$_669e[670], 0);
                        a.setAttribute(_$_669e[671], 100);
                        a.onblur = function() {
                            k.closeEditor(d, !0)
                        };
                        a.focus();
                        a.value = e
                    };
                b.closeEditor = function(d, e, g, f, k) {
                    return e ? a.value : _$_669e[3]
                };
                return b
            }();
        h.rating = function() {
            var b = {},
                a = null;
            b.createCell = function(d, e, g, f, k, l) {
                return b.setCell(d, e, l)
            };
            b.destroyCell = function(d) {
                d.classList.remove(_$_669e[672])
            };
            b.updateCell = b.createCell;
            b.openEditor = function(d, e, g, f, k, l) {
                a = h.createEditor(_$_669e[269],
                    d, e, k);
                a.type = _$_669e[22];
                a.setAttribute(_$_669e[670], 0);
                a.setAttribute(_$_669e[671], 5);
                a.onblur = function() {
                    k.closeEditor(d, !0)
                };
                a.focus();
                a.value = e
            };
            b.closeEditor = function(d, e, g, f, k, l) {
                return e ? a.value : _$_669e[3]
            };
            b.setCell = function(d, e, g) {
                e = parseInt(e);
                if (5 < e) e = 5;
                else if (!e || 0 > e) e = 0;
                if (d) {
                    g = g && g.color ? g.color : _$_669e[649];
                    var f = document.createElement(_$_669e[32]);
                    f.setAttribute(_$_669e[315], parseInt(e) + _$_669e[673]);
                    f.classList.add(_$_669e[307]);
                    for (var k = 0; k < e; k++) {
                        var l = document.createElement(_$_669e[191]);
                        l.className = _$_669e[192];
                        l.style.color = g;
                        l.innerHTML = _$_669e[674];
                        f.appendChild(l)
                    }
                    d.innerHTML = _$_669e[3];
                    d.className = _$_669e[672];
                    d.appendChild(f)
                }
                return e
            };
            return b
        }();
        h.email = function() {
            var b = {
                createCell: function(a, d, e, g, f, k) {
                    a && (a.children && a.children[0] && a.children[0].tagName == _$_669e[632] ? e = a.children[0] : (e = document.createElement(_$_669e[245]), a.innerText = _$_669e[3], a.appendChild(e), k.options && k.options.url && e.setAttribute(_$_669e[675], k.options.url)), e.innerText = d)
                }
            };
            b.updateCell = b.createCell;
            b.openEditor = function(a, d, e, g, f, k) {
                f.parent.input.classList.add(_$_669e[624]);
                f.parent.input.onblur = function() {
                    f.closeEditor(a, !0)
                };
                f.parent.input.innerText = d;
                jSuites.focus(f.parent.input)
            };
            b.closeEditor = function(a, d, e, g, f) {
                f.parent.input.classList.remove(_$_669e[624]);
                return d ? f.parent.input.innerText.replace(new RegExp(/\n/, _$_669e[14]), _$_669e[3]) : _$_669e[3]
            };
            return b
        }();
        h.url = h.email;
        h.image = function() {
            var b = {},
                a = null;
            b.createCell = function(d, e, g, f, k, l) {
                if (d)
                    if (e)
                        if (d.children && d.children[0] &&
                            d.children[0].tagName == _$_669e[676]) d.children[0].src = e;
                        else {
                            var n = document.createElement(_$_669e[49]);
                            l.render == _$_669e[677] && n.classList.add(_$_669e[677]);
                            l.options && (n.style.maxWidth = _$_669e[678], l.options.absolute && (d.classList.add(_$_669e[679]), n.setAttribute(_$_669e[680], -1), n.classList.add(_$_669e[270]), n.classList.add(_$_669e[681]), n.settings = function() {
                                    k.openEditor(d, e, g, f, k, l)
                                }, n.dblclick = function() {
                                    k.openEditor(d, !0)
                                }, n[_$_669e[682]] = function() {
                                    k.setValueFromCoords(g, f, _$_669e[3])
                                }, n.refresh =
                                function(p) {
                                    if (!p) {
                                        p = parseInt(n.style.top) || 0;
                                        var q = parseInt(n.style.left) || 0,
                                            v = parseInt(n.style.width) || 40,
                                            w = parseInt(n.style.height) || 40,
                                            x = d.getAttribute(_$_669e[63]),
                                            y = d.getAttribute(_$_669e[59]);
                                        k.updateProperty([{
                                            x,
                                            y,
                                            value: {
                                                top: p,
                                                left: q,
                                                width: v,
                                                height: w
                                            }
                                        }], null, null, !0)
                                    }
                                }), l.options.width && (n.style.width = parseInt(l.options.width) + _$_669e[57]), l.options.height && (n.style.height = parseInt(l.options.height) + _$_669e[57]), l.top && (n.style.top = parseInt(l.top) + _$_669e[57]), l.left && (n.style.left = parseInt(l.left) +
                                _$_669e[57]), l.width && (n.style.width = parseInt(l.width) + _$_669e[57]), l.height && (n.style.height = parseInt(l.height) + _$_669e[57]));
                            n.src = e;
                            d.innerHTML = _$_669e[3];
                            d.appendChild(n)
                        }
                else d.innerHTML = _$_669e[3]
            };
            b.updateCell = b.createCell;
            b.openEditor = function(d, e, g, f, k, l) {
                (l && l.options ? l.options : {}).value = e;
                a = h.createDialog(d, e, g, f, k, l);
                e && (d = document.createElement(_$_669e[49]), d.src = e, d.classList.add(_$_669e[683]), d.style.width = _$_669e[678], a.appendChild(d));
                jSuites.image(a, l)
            };
            b.closeEditor = function(d,
                e, g, f, k, l) {
                if (e && (d = a.children[0])) return d.tagName == _$_669e[676] ? d.src : _$_669e[3]
            };
            b[_$_669e[56]] = function(d, e) {
                return _$_669e[684] + e + _$_669e[685]
            };
            return b
        }();
        h.html = function() {
            var b = {},
                a = null;
            b.createCell = function(d, e) {
                d.classList.add(_$_669e[686]);
                var g = document.createElement(_$_669e[32]);
                g.innerHTML = e;
                d.appendChild(g)
            };
            b.updateCell = function(d, e) {
                d && (d.firstChild.innerHTML = e)
            };
            b.destroyCell = function(d) {
                d.classList.remove(_$_669e[686])
            };
            b.openEditor = function(d, e, g, f, k, l) {
                var n = l && l.options ? l.options : {};
                n.focus = !0;
                n.value = e;
                n.border = !1;
                n.height = _$_669e[651];
                e = h.createDialog(d, e, g, f, k, l);
                a = jSuites.editor(e, n);
                a.close = function() {
                    k.closeEditor(d, !0)
                };
                jSuites.tracking(a, !0)
            };
            b.closeEditor = function(d, e, g, f, k, l) {
                jSuites.tracking(a, !1);
                d = a.getData();
                if (e) return d
            };
            return b
        }();
        h.hidden = function() {
            return {
                createCell: function(b, a, d, e, g) {
                    b.innerText = a
                },
                updateCell: function(b, a, d, e, g) {
                    b && (b.innerText = a)
                },
                openEditor: function(b, a, d, e, g) {
                    return !1
                },
                closeEditor: function(b, a) {
                    return !1
                }
            }
        }();
        h.tags = function() {
            var b = {},
                a = null;
            b.createCell = function(d, e) {
                d.classList.add(_$_669e[687]);
                b.updateCell(d, e)
            };
            b.updateCell = function(d, e) {
                d && (d.innerHTML = e)
            };
            b.destroyCell = function(d) {
                d.classList.remove(_$_669e[687])
            };
            b.openEditor = function(d, e, g, f, k, l) {
                var n = l && l.options ? l.options : {};
                n.value = e;
                a = h.createDialog(d, e, g, f, k, l);
                d = document.createElement(_$_669e[32]);
                d.style.margin = _$_669e[34];
                d.style.marginRight = _$_669e[688];
                a.appendChild(d);
                a = jSuites.tags(d, n);
                d.focus();
                jSuites.focus(d)
            };
            b.closeEditor = function(d, e, g, f, k, l) {
                d =
                    a.getValue();
                if (e) return d
            };
            return b
        }();
        h.record = function() {
            var b = {},
                a = null;
            b.createCell = function(d, e, g, f, k, l) {
                d.classList.add(_$_669e[637]);
                var n;
                (n = k.parent.getWorksheetInstance(l.worksheetId)) ? n.rows ? b.updateCell(d, e, g, f, k, l) : (d.innerText = _$_669e[293], k.parent.queue.push([k, g, f])): (d.innerText = _$_669e[293], k.parent.queue.push([k, g, f]))
            };
            b.destroyCell = function(d) {
                d.classList.remove(_$_669e[637])
            };
            b.updateCell = function(d, e, g, f, k, l) {
                d && (typeof l.worksheetId == _$_669e[21] ? d.innerText = _$_669e[293] : d.innerHTML =
                    b.getValue(d, e, g, f, k, l))
            };
            b.openEditor = function(d, e, g, f, k, l) {
                var n = l.delimiter || _$_669e[262];
                l = Object.create(l);
                l.options || (l.options = {});
                typeof l.filter == _$_669e[110] && (l.source = l.filter(k.el, d, g, f, l.source));
                typeof l.filterOptions == _$_669e[110] && (l = l.filterOptions(k, d, g, f, e, l));
                g = l.options;
                g.data = b.getSource(k, l);
                c(l, _$_669e[654]) && (g.multiple = !0);
                c(l, _$_669e[70]) && (g.autocomplete = !0);
                g.format = !0;
                g.opened = !0;
                g.width = d.offsetWidth - 2;
                g.onclose = function() {
                    k.closeEditor(d, !0)
                };
                c(l, _$_669e[654]) ? e &&
                    (g.value = (_$_669e[3] + e).split(n)) : g.value = e;
                a = h.createEditor(_$_669e[32], d, e, k);
                a = jSuites.dropdown(a, g)
            };
            b.closeEditor = function(d, e, g, f, k, l) {
                d = l.delimiter || _$_669e[262];
                g = a.getValue(!0);
                a.close(!0);
                if (e) return g.join(d)
            };
            b.getValue = function(d, e, g, f, k, l) {
                d = [];
                if (typeof l.worksheetId == _$_669e[25] && (k = k.parent.getWorksheetInstance(l.worksheetId))) {
                    g = l.worksheetColumn ? l.worksheetColumn : 0;
                    0 <= l.worksheetImage && (g = l.worksheetImage);
                    if (e == _$_669e[3] || 0 == e || null == e || void 0 == e) return _$_669e[3];
                    e = (_$_669e[3] +
                        e).split(_$_669e[262]);
                    for (f = 0; f < e.length; f++) {
                        var n = e[f].trim();
                        var p = k.name(g);
                        n = k.getRowById(n);
                        p = !1 === n ? _$_669e[689] : n[p];
                        0 <= l.worksheetImage && (p = _$_669e[690] + p + _$_669e[691]);
                        d.push(p)
                    }
                    return 0 <= l.worksheetImage ? d.join(_$_669e[3]) : d.join(_$_669e[353])
                }
                return _$_669e[293]
            };
            b.getSource = function(d, e) {
                var g = e.worksheetColumn ? e.worksheetColumn : 0,
                    f = null;
                typeof e.worksheetId == _$_669e[25] && (f = d.parent.getWorksheetInstance(e.worksheetId));
                if (f) {
                    var k, l = [];
                    for (d = 0; d < f.rows.length; d++)
                        if (k = f.value(g, d)) k = {
                            id: f.getRowId(d),
                            name: k
                        }, void 0 != e.worksheetImage && (k.image = f.value(e.worksheetImage, d)), l.push(k)
                } else console.error(_$_669e[692] + e.worksheetId);
                return l
            };
            b[_$_669e[56]] = function(d, e) {
                for (var g = 0; g < d.source.length; g++)
                    if (d.source[g].id == e) return d.render == _$_669e[305] ? d.source[g].color : d.render == _$_669e[655] ? d.source[g].image : d.source[g].name
            };
            return b
        }();
        return h
    }();
    typeof jQuery !== _$_669e[21] && function(h) {
        h.fn.jspreadsheet = h.fn.jexcel = function(c) {
            var b = h(this)[_$_669e[56]](0);
            if (b.jspreadsheet) {
                if (typeof arguments[0] ==
                    _$_669e[22]) var a = arguments[0],
                    d = 2;
                else a = 0, d = 1;
                return b.jspreadsheet[a][c].apply(b.jspreadsheet[a], Array.prototype.slice.call(arguments, d))
            }
            return r(h(this)[_$_669e[56]](0), arguments[0])
        }
    }(jQuery);
    return r
});