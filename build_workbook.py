"""Write the Tableau workbook (ewallet_review_insights.twb) as XML.

Building the .twb in code keeps every sheet, calculation, colour and dashboard action
reviewable in git and reproducible. Open the file in Tableau, check it, then save.
"""
import uuid
from pathlib import Path
from xml.dom import minidom
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent
DIR = HERE.as_posix()
R = "federated.ewalletreviews01"   # review data source
M = "federated.ewalletmarket01"    # Bank Indonesia market data source
P = "Parameters"
NL = "Æ&#10;"                      # line break inside a Tableau formatted-text run
APPS = ["DANA", "GoPay", "OVO", "ShopeePay", "LinkAja"]
VERSIONS = ["2.141.0", "2.142.0", "2.143.0", "2.144.0", "2.145.0"]
NO_THEME = "(no complaint theme)"


def q(s):
    return escape(s, {"'": "&apos;", '"': "&quot;"})


def uid():
    return "{" + str(uuid.uuid4()).upper() + "}"


# ---------------------------------------------------------------- data sources
REVIEW_COLS = [  # name, datatype, role, type
    ("Review ID", "string", "dimension", "nominal"), ("App", "string", "dimension", "nominal"),
    ("Peer Group", "string", "dimension", "nominal"), ("Lens", "string", "dimension", "nominal"),
    ("Date", "date", "dimension", "ordinal"), ("Rating", "integer", "dimension", "ordinal"),
    ("Version", "string", "dimension", "nominal"), ("Theme", "string", "dimension", "nominal"),
    ("Theme Count", "integer", "measure", "quantitative"), ("Share Weight", "real", "measure", "quantitative"),
    ("First Row", "integer", "measure", "quantitative"), ("One Star", "integer", "measure", "quantitative"),
    ("Negative", "integer", "measure", "quantitative"), ("Helpful Votes", "integer", "measure", "quantitative"),
    ("Snippet", "string", "dimension", "nominal"),
]
MARKET_COLS = [("Metric", "string", "dimension", "nominal"), ("Period Type", "string", "dimension", "nominal"),
               ("Period", "date", "dimension", "ordinal"), ("Value", "real", "measure", "quantitative")]

# name: (caption, datatype, role, type, format, formula)
PARAM_APP = f"[{P}].[Parameter 1]"
PARAM_LENS = f"[{P}].[Parameter 2]"
REVIEW_CALCS = {
    "Calc_AppFilter": ("App in selection", "boolean", "dimension", "nominal", None,
                       f'{PARAM_APP} = "All apps" OR [App] = {PARAM_APP}'),
    "Calc_LensFilter": ("Lens in selection", "boolean", "dimension", "nominal", None, f"[Lens] = {PARAM_LENS}"),
    "Calc_IsTheme": ("Has complaint theme", "boolean", "dimension", "nominal", None, f'[Theme] <> "{NO_THEME}"'),
    "Calc_Share": ("% of Reviews", "real", "measure", "quantitative", "p0.0%",
                   "SUM([Share Weight]) / COUNTD([App])"),
    "Calc_Severity": ("1-Star Share of Theme", "real", "measure", "quantitative", "p0%", "AVG([One Star])"),
    "Calc_Reviews": ("Reviews", "integer", "measure", "quantitative", "n#,##0;-#,##0", "COUNTD([Review ID])"),
    "Calc_AvgRating": ("Avg Rating", "real", "measure", "quantitative", "n#,##0.00;-#,##0.00",
                       "SUM([Rating] * [First Row]) / SUM([First Row])"),
    "Calc_OneStar": ("1-Star Reviews", "real", "measure", "quantitative", "p0%",
                     "SUM([One Star] * [First Row]) / SUM([First Row])"),
    "Calc_PerDay": ("Reviews per Day", "real", "measure", "quantitative", "n#,##0.0;-#,##0.0",
                    "COUNTD([Review ID]) / (DATEDIFF('day', MIN([Date]), MAX([Date])) + 1)"),
    "Calc_Days": ("Days Covered", "integer", "measure", "quantitative", "n#,##0;-#,##0",
                  "DATEDIFF('day', MIN([Date]), MAX([Date])) + 1"),
    "Calc_MixShare": ("Share of App Reviews", "real", "measure", "quantitative", "p0%",
                      "SUM([First Row]) / MIN({FIXED [App], [Lens] : SUM([First Row])})"),
    "Calc_DanaShare": ("DANA %", "real", "measure", "quantitative", "p0.0%",
                       'SUM(IF [App] = "DANA" THEN [Share Weight] END)'),
    "Calc_PeerShare": ("Peer Avg %", "real", "measure", "quantitative", "p0.0%",
                       'SUM(IF [App] <> "DANA" THEN [Share Weight] END) / 4'),
    "Calc_Gap": ("DANA Gap vs Peers (pp)", "real", "measure", "quantitative", "n+#,##0.0;-#,##0.0",
                 "([Calc_DanaShare] - [Calc_PeerShare]) * 100"),
    "Calc_GapColor": ("Gap Direction", "string", "measure", "nominal", None,
                      'IF [Calc_Gap] > 0 THEN "DANA worse than peers" ELSE "DANA better than peers" END'),
    "Calc_Stars": ("Stars", "string", "dimension", "nominal", None, 'STR([Rating]) + " star"'),
}
MARKET_CALCS = {
    "Calc_V2025": ("2025", "real", "measure", "quantitative", "n#,##0;-#,##0",
                   'SUM(IF [Period Type] = "Annual" AND YEAR([Period]) = 2025 THEN [Value] END)'),
    "Calc_Growth": ("Growth since 2021", "real", "measure", "quantitative", "n#,##0.0;-#,##0.0",
                    '[Calc_V2025] / SUM(IF [Period Type] = "Annual" AND YEAR([Period]) = 2021 THEN [Value] END)'),
}

THEME_COLORS = {
    "Failed transfer": "#e15759", "Customer service": "#f28e2b", "Login, OTP & KYC": "#4e79a7",
    "App bugs & speed": "#76b7b2", "Balance lost": "#b07aa1", "Paylater & bills": "#edc948",
    "Fees & charges": "#59a14f", "Hard to use": "#9c755f", "Security & fraud": "#ff9da7",
    "Ads & promos": "#bab0ac", NO_THEME: "#d9d9d9",
}
RATING_COLORS = {1: "#d7301f", 2: "#fc8d59", 3: "#d9d9d9", 4: "#91bfdb", 5: "#4575b4"}
PEER_COLORS = {"DANA": "#118dff", "Peers": "#b8bcc2"}
GAP_COLORS = {"DANA worse than peers": "#e15759", "DANA better than peers": "#4e9f94"}


def column_defs(cols, calcs):
    out = []
    for name, dt, role, typ in cols:
        out.append(f"<column caption='{q(name)}' datatype='{dt}' name='[{q(name)}]' role='{role}' type='{typ}' />")
    for name, (cap, dt, role, typ, fmt, formula) in calcs.items():
        f = f" default-format='{q(fmt)}'" if fmt else ""
        out.append(f"<column caption='{q(cap)}' datatype='{dt}'{f} name='[{name}]' role='{role}' type='{typ}'>"
                   f"<calculation class='tableau' formula='{q(formula)}' /></column>")
    return "\n".join(out)


def datasource(ds, caption, filename, cols, calcs, style="", ds_instances=""):
    rel_cols = "\n".join(f"<column datatype='{dt}' name='{q(n)}' ordinal='{i}' />" for i, (n, dt, _, _) in enumerate(cols))
    conn = f"textscan.{ds.split('.')[1]}"
    table = filename.replace(".", "#")
    return f"""<datasource caption='{caption}' inline='true' name='{ds}' version='18.1'>
<connection class='federated'>
<named-connections><named-connection caption='{caption}' name='{conn}'>
<connection class='textscan' directory='{q(DIR)}' filename='{filename}' password='' server='' workgroup-auth-mode='as-is' />
</named-connection></named-connections>
<relation connection='{conn}' name='{filename}' table='[{table}]' type='table'>
<columns character-set='UTF-8' header='yes' locale='en_US' separator=','>
{rel_cols}
</columns></relation>
</connection>
<aliases enabled='yes' />
{column_defs(cols, calcs)}
{ds_instances}
<extract count='-1' enabled='true' object-id='' units='records' user-specific='false'>
<connection access_mode='readonly' author-locale='en_US' class='hyper' dbname='C:/Users/Emelio Exaudi/Documents/My Tableau Repository/Datasources/{caption}.hyper' default-settings='hyper' schema='Extract' sslmode='' tablename='Extract' update-time='09/26/2026 09:50:00 PM' username='tableau_internal_user'>
<relation name='Extract' table='[Extract].[Extract]' type='table' /></connection></extract>
<layout dim-ordering='alphabetic' measure-ordering='alphabetic' show-structure='true' />
{style}
</datasource>"""


def color_map(field, mapping, quote=True):
    maps = "".join(f"<map to='{c}'><bucket>{'&quot;' + q(k) + '&quot;' if quote else k}</bucket></map>"
                   for k, c in mapping.items())
    return f"<encoding attr='color' field='{field}' type='palette'>{maps}</encoding>"


REVIEW_STYLE = ("<style><style-rule element='mark'>"
                + color_map("[none:Theme:nk]", THEME_COLORS)
                + color_map("[none:Rating:ok]", RATING_COLORS, quote=False)
                + color_map("[none:Peer Group:nk]", {"DANA": "#118dff", "Peers": "#b8bcc2"})
                + color_map("[usr:Calc_GapColor:nk]", {"DANA worse than peers": "#e15759",
                                                        "DANA better than peers": "#4e9f94"})
                + "</style-rule></style>")


# colour maps only take effect when the data source also declares the matching column instance
REVIEW_INSTANCES = (
    "<column-instance column='[Theme]' derivation='None' name='[none:Theme:nk]' pivot='key' type='nominal' />"
    "<column-instance column='[Rating]' derivation='None' name='[none:Rating:ok]' pivot='key' type='ordinal' />"
    "<column-instance column='[Peer Group]' derivation='None' name='[none:Peer Group:nk]' pivot='key' type='nominal' />"
    "<column-instance column='[Calc_GapColor]' derivation='User' name='[usr:Calc_GapColor:nk]' pivot='key' type='nominal' />")


def parameters():
    apps = "".join(f"<member value='&quot;{a}&quot;' />" for a in ["All apps"] + APPS)
    lens = "".join(f"<member value='&quot;{a}&quot;' />" for a in ["Latest 500", "Most helpful"])
    return f"""<datasource hasconnection='false' inline='true' name='{P}' version='18.1'>
<aliases enabled='yes' />
<column caption='App' datatype='string' name='[Parameter 1]' param-domain-type='list' role='measure' type='nominal' value='&quot;All apps&quot;'>
<calculation class='tableau' formula='&quot;All apps&quot;' /><members>{apps}</members></column>
<column caption='Reviews' datatype='string' name='[Parameter 2]' param-domain-type='list' role='measure' type='nominal' value='&quot;Latest 500&quot;'>
<calculation class='tableau' formula='&quot;Latest 500&quot;' /><members>{lens}</members></column>
</datasource>"""


# ---------------------------------------------------------------- worksheets
DERIV = {"none": "None", "sum": "Sum", "usr": "User", "yr": "Year", "max": "Max", "mdy": "MDY", "avg": "Avg"}


def ref(ds, deriv, field, kind):
    return f"[{ds}].[{deriv}:{field}:{kind}]"


def instances(refs):
    seen, out = set(), []
    for r in refs:
        if r in seen:
            continue
        seen.add(r)
        ds, inst = r.split("].[", 1)
        deriv, field, kind = inst.rstrip("]").split(":")
        typ = {"nk": "nominal", "ok": "ordinal", "qk": "quantitative"}[kind]
        out.append((ds.lstrip("["), f"<column-instance column='[{field}]' derivation='{DERIV[deriv]}' "
                                    f"name='[{deriv}:{field}:{kind}]' pivot='key' type='{typ}' />"))
    return out


def run(text, size=None, bold=False, color=None, raw=False):
    attrs = (f" bold='true'" if bold else "") + (f" fontcolor='{color}'" if color else "") + (f" fontsize='{size}'" if size else "")
    return f"<run{attrs}>{text if raw else q(text)}</run>"


def field_run(r, **kw):
    return run(f"&lt;{q(r)}&gt;", raw=True, **kw)


def sheet(name, ds, rows="", cols="", mark="Automatic", enc=None, filters=(), sort=None, tooltip=None,
          label=None, show_labels=False, params=False, extra_refs=(), ws_style="", title=None, mark_size=None,
          colors=(), hide=(), cull=True):
    enc = enc or {}
    refs = [x for x in (rows.replace("(", "").replace(")", "").split(" / ") + cols.split(" / ")) if x.strip()]
    for v in enc.values():
        refs += v if isinstance(v, list) else [v]
    refs += [f[0] for f in filters] + list(extra_refs)
    if sort:
        refs += [sort[0], sort[1]]
    inst = instances(refs)
    cols_defs = column_defs(REVIEW_COLS, REVIEW_CALCS) if ds == R else column_defs(MARKET_COLS, MARKET_CALCS)
    dep = f"<datasource-dependencies datasource='{ds}'>{cols_defs}{''.join(i for d, i in inst if d == ds)}</datasource-dependencies>"
    pdep = ""
    if params:
        pdep = (f"<datasource-dependencies datasource='{P}'>"
                "<column caption='App' datatype='string' name='[Parameter 1]' param-domain-type='list' role='measure' type='nominal' value='&quot;All apps&quot;'><calculation class='tableau' formula='&quot;All apps&quot;' /></column>"
                "<column caption='Reviews' datatype='string' name='[Parameter 2]' param-domain-type='list' role='measure' type='nominal' value='&quot;Latest 500&quot;'><calculation class='tableau' formula='&quot;Latest 500&quot;' /></column>"
                "</datasource-dependencies>")
    fx = []
    for col, members in filters:
        lvl = col.split("].[", 1)[1][:-1]
        if len(members) == 1:
            g = f"<groupfilter function='member' level='[{lvl}]' member='{q(members[0])}' user:ui-domain='database' user:ui-enumeration='inclusive' user:ui-marker='enumerate' />"
        else:
            kids = "".join(f"<groupfilter function='member' level='[{lvl}]' member='{q(m)}' />" for m in members)
            g = f"<groupfilter function='union' user:ui-domain='database' user:ui-enumeration='inclusive' user:ui-marker='enumerate'>{kids}</groupfilter>"
        fx.append(f"<filter class='categorical' column='{col}'>{g}</filter>")
    sorts = ""
    if sort:
        sorts = f"<computed-sort column='{sort[0]}' direction='{sort[2]}' using='{sort[1]}' />"
    slices = "".join(f"<column>{c}</column>" for c, _ in filters)
    encs = ""
    for k, v in enc.items():
        for r in (v if isinstance(v, list) else [v]):
            encs += f"<{k} column='{r}' />"
    lab = f"<customized-label><formatted-text>{label}</formatted-text></customized-label>" if label else ""
    tip = f"<customized-tooltip><formatted-text>{tooltip}</formatted-text></customized-tooltip>" if tooltip else ""
    mstyle = ("<style><style-rule element='mark'><format attr='mark-labels-show' value='true' />"
              f"<format attr='mark-labels-cull' value='{str(cull).lower()}' /></style-rule></style>") if show_labels else ""
    # categorical colours are set per worksheet (data-source level maps were ignored)
    if colors:
        ws_style += "<style-rule element='mark'>" + "".join(color_map(f, mp, qt) for f, mp, qt in colors) + "</style-rule>"
    if hide:
        ws_style += "<style-rule element='label'>" + "".join(
            f"<format attr='display' field='{h}' value='false' />" for h in hide) + "</style-rule>"
    size = f"<size column='{mark_size}' />" if mark_size else ""
    ttl = f"<layout-options><title><formatted-text>{run(title or name, size=11, bold=True)}</formatted-text></title></layout-options>"
    dss = f"<datasource name='{ds}' />" + (f"<datasource name='{P}' />" if params else "")
    return f"""<worksheet name='{q(name)}'>{ttl}
<table><view><datasources>{dss}</datasources>{pdep}{dep}{sorts}{''.join(fx)}
{'<slices>' + slices + '</slices>' if slices else ''}<aggregation value='true' /></view>
<style>{ws_style}</style>
<panes><pane selection-relaxation-option='selection-relaxation-allow'><view><breakdown value='auto' /></view>
<mark class='{mark}' /><encodings>{encs}{size}</encodings>{tip}{lab}{mstyle}</pane></panes>
<rows>{rows}</rows><cols>{cols}</cols></table>
<simple-id uuid='{uid()}' /></worksheet>"""


r = lambda d, f, k: ref(R, d, f, k)
m = lambda d, f, k: ref(M, d, f, k)
APPF = (r("none", "Calc_AppFilter", "nk"), ["true"])
LENSF = (r("none", "Calc_LensFilter", "nk"), ["true"])
THEMEF = (r("none", "Calc_IsTheme", "nk"), ["true"])
LATEST = (r("none", "Lens", "nk"), ['"Latest 500"'])
SHARE, SEV, REV = r("usr", "Calc_Share", "qk"), r("usr", "Calc_Severity", "qk"), r("usr", "Calc_Reviews", "qk")
AVG, ONE, DAYS = r("usr", "Calc_AvgRating", "qk"), r("usr", "Calc_OneStar", "qk"), r("usr", "Calc_Days", "qk")
PERDAY = r("usr", "Calc_PerDay", "qk")
THEME, APP, RATING = r("none", "Theme", "nk"), r("none", "App", "nk"), r("none", "Rating", "ok")
PEER = r("none", "Peer Group", "nk")

S = {}
S["market_value"] = sheet(
    "E-money transaction value", M, rows=m("sum", "Value", "qk"), cols=m("yr", "Period", "ok"), mark="Bar",
    filters=[(m("none", "Period Type", "nk"), ['"Annual"']), (m("none", "Metric", "nk"), ['"Transaction value (Rp trillion)"'])],
    enc={"text": m("sum", "Value", "qk")}, show_labels=True,
    title="E-money transaction value per year (Rp trillion)")
S["scorecard"] = sheet(
    "Average rating per app", R, rows=APP, cols=AVG, mark="Bar", params=True, filters=[LENSF],
    enc={"color": PEER, "text": AVG, "tooltip": [ONE, REV, PERDAY]}, sort=(APP, AVG, "DESC"), show_labels=True,
    colors=[(PEER, PEER_COLORS, True)],
    tooltip=field_run(APP, bold=True) + run(NL, raw=True) + run("Average rating: ") + field_run(AVG, bold=True)
    + run(NL, raw=True) + run("1-star reviews: ") + field_run(ONE) + run(NL, raw=True) + run("Reviews in lens: ") + field_run(REV),
    title="Average App Store rating (1-5)")
S["mix"] = sheet(
    "Star-rating mix per app", R, rows=APP, cols=r("usr", "Calc_MixShare", "qk"), mark="Bar", params=True,
    filters=[LENSF], enc={"color": RATING, "tooltip": [REV]}, sort=(APP, AVG, "DESC"), extra_refs=[AVG],
    colors=[(RATING, RATING_COLORS, False)],
    tooltip=field_run(APP, bold=True) + run(": ") + field_run(r("usr", "Calc_MixShare", "qk"), bold=True)
    + run(" of reviews gave ") + field_run(RATING) + run(" star(s)"),
    title="Star-rating mix (red = 1 star, blue = 5 stars)")
S["velocity"] = sheet(
    "Reviews per day", R, rows=APP, cols=PERDAY, mark="Bar", filters=[LATEST],
    enc={"color": PEER, "text": PERDAY, "tooltip": [DAYS]}, sort=(APP, PERDAY, "DESC"), show_labels=True,
    colors=[(PEER, PEER_COLORS, True)],
    tooltip=field_run(APP, bold=True) + run(NL, raw=True) + run("Latest 500 reviews span ") + field_run(DAYS)
    + run(" days = ") + field_run(PERDAY, bold=True) + run(" reviews per day"),
    title="Review velocity: new reviews per day")
S["kpi_reviews"] = sheet(
    "Review KPIs", R, mark="Text", params=True, filters=[APPF, LENSF],
    enc={"text": [REV, AVG, ONE]},
    label=field_run(REV, size=18, bold=True) + run(" reviews   ", size=10, color="#6b7785")
    + field_run(AVG, size=18, bold=True) + run(" avg stars   ", size=10, color="#6b7785")
    + field_run(ONE, size=18, bold=True, color="#d7301f") + run(" gave 1 star", size=10, color="#6b7785"),
    show_labels=True, title="Selection at a glance")
S["heatmap"] = sheet(
    "Complaint heatmap", R, rows=THEME, cols=APP, mark="Square", params=True, filters=[APPF, LENSF, THEMEF],
    enc={"color": SHARE, "text": SHARE, "tooltip": [SEV, REV]}, sort=(THEME, SHARE, "DESC"), show_labels=True,
    ws_style=f"<style-rule element='mark'><encoding attr='color' field='{SHARE}' palette='orange_10_0' type='interpolated' /></style-rule>",
    tooltip=field_run(THEME, bold=True) + run(" - ") + field_run(APP, bold=True) + run(NL, raw=True)
    + field_run(SHARE, bold=True) + run(" of reviews mention it (") + field_run(REV) + run(" reviews)") + run(NL, raw=True)
    + field_run(SEV) + run(" of them gave 1 star") + run(NL, raw=True) + run("Click to read these reviews", color="#6b7785"),
    title="% of reviews mentioning each complaint theme (click a cell)")
S["drill_rating"] = sheet(
    "Star ratings of selection", R, rows=REV, cols=RATING, mark="Bar", params=True, filters=[APPF, LENSF],
    enc={"color": RATING, "text": REV}, show_labels=True, colors=[(RATING, RATING_COLORS, False)],
    title="Star ratings of the selected reviews")
S["examples"] = sheet(
    "What users wrote", R, rows=f"({APP} / {RATING} / {r('none', 'Review ID', 'nk')})", mark="Text", params=True,
    filters=[APPF, LENSF], enc={"text": r("none", "Snippet", "nk")}, extra_refs=[r("max", "Helpful Votes", "qk")],
    sort=(r("none", "Review ID", "nk"), r("max", "Helpful Votes", "qk"), "DESC"), hide=[r("none", "Review ID", "nk")],
    label=run("&lt;" + q(r("none", "Snippet", "nk")) + "&gt;", raw=True, size=9), show_labels=True, cull=False,
    ws_style="<style-rule element='cell'><format attr='height' value='30' /><format attr='width' value='400' /></style-rule>",
    tooltip=field_run(r("none", "Snippet", "nk")) + run(NL, raw=True) + run("Helpful votes: ", color="#6b7785")
    + field_run(r("max", "Helpful Votes", "qk"), color="#6b7785"),
    title="What users wrote (sorted by helpful votes)")
S["priority"] = sheet(
    "Priority matrix", R, rows=SEV, cols=SHARE, mark="Circle", params=True, filters=[APPF, LENSF, THEMEF],
    enc={"color": THEME, "text": THEME, "tooltip": [REV]}, mark_size=REV, show_labels=True,
    colors=[(THEME, THEME_COLORS, True)],
    tooltip=field_run(THEME, bold=True) + run(NL, raw=True) + run("How common: ") + field_run(SHARE, bold=True)
    + run(" of reviews") + run(NL, raw=True) + run("How severe: ") + field_run(SEV, bold=True) + run(" of those gave 1 star"),
    title="Priority matrix: how common (x) vs how severe (y)")
S["gap"] = sheet(
    "DANA vs peers", R, rows=THEME, cols=r("usr", "Calc_Gap", "qk"), mark="Bar", params=True, filters=[LENSF, THEMEF],
    enc={"color": r("usr", "Calc_GapColor", "nk"), "text": r("usr", "Calc_Gap", "qk"),
         "tooltip": [r("usr", "Calc_DanaShare", "qk"), r("usr", "Calc_PeerShare", "qk")]},
    sort=(THEME, r("usr", "Calc_Gap", "qk"), "DESC"), show_labels=True,
    colors=[(r("usr", "Calc_GapColor", "nk"), GAP_COLORS, True)],
    tooltip=field_run(THEME, bold=True) + run(NL, raw=True) + run("DANA: ") + field_run(r("usr", "Calc_DanaShare", "qk"), bold=True)
    + run(" of reviews   Peer average: ") + field_run(r("usr", "Calc_PeerShare", "qk")) + run(NL, raw=True)
    + run("Gap: ") + field_run(r("usr", "Calc_Gap", "qk"), bold=True) + run(" percentage points"),
    title="DANA vs peer average (gap in percentage points)")
S["versions"] = sheet(
    "DANA by app version", R, rows=AVG, cols=r("none", "Version", "nk"), mark="Bar",
    filters=[(APP, ['"DANA"']), LATEST, (r("none", "Version", "nk"), [f'"{v}"' for v in VERSIONS])],
    enc={"color": ONE, "text": AVG, "tooltip": [REV, ONE]}, show_labels=True,
    ws_style=f"<style-rule element='mark'><encoding attr='color' field='{ONE}' palette='red_10_0' type='interpolated' /></style-rule>",
    tooltip=run("DANA ") + field_run(r("none", "Version", "nk"), bold=True) + run(NL, raw=True) + run("Average rating: ")
    + field_run(AVG, bold=True) + run(NL, raw=True) + run("1-star share: ") + field_run(ONE) + run(NL, raw=True)
    + run("Reviews: ") + field_run(REV),
    title="DANA rating by app version (Sept 2026)")


# ---------------------------------------------------------------- dashboards
STYLE = ("<zone-style><format attr='border-color' value='#000000' /><format attr='border-style' value='none' />"
         "<format attr='border-width' value='0' /><format attr='margin' value='4' /></zone-style>")
_ids = iter(range(1, 999))


def z_sheet(name, x, y, w, h, show_title=True):
    t = "" if show_title else " show-title='false'"
    return f"<zone h='{h}' id='{next(_ids)}' name='{q(name)}'{t} w='{w}' x='{x}' y='{y}'>{STYLE}</zone>"


def z_text(runs, x, y, w, h, bg=None):
    b = f"<format attr='background-color' value='{bg}' />" if bg else ""
    st = STYLE.replace("</zone-style>", b + "</zone-style>")
    return f"<zone h='{h}' id='{next(_ids)}' type-v2='text' w='{w}' x='{x}' y='{y}'><formatted-text>{runs}</formatted-text>{st}</zone>"


def z_param(param, x, y, w, h):
    return f"<zone h='{h}' id='{next(_ids)}' mode='compact' param='{param}' type-v2='paramctrl' w='{w}' x='{x}' y='{y}'>{STYLE}</zone>"


def dashboard(name, zones):
    return f"""<dashboard enable-sort-zone-taborder='true' name='{q(name)}'>
<style />
<size maxheight='800' maxwidth='1300' minheight='800' minwidth='1300' sizing-mode='fixed' />
<zones><zone h='100000' id='{next(_ids)}' type-v2='layout-basic' w='100000' x='0' y='0'>
{''.join(zones)}
<zone-style><format attr='border-color' value='#000000' /><format attr='border-style' value='none' /><format attr='border-width' value='0' /><format attr='margin' value='8' /></zone-style>
</zone></zones>
<simple-id uuid='{uid()}' /></dashboard>"""


def kpi(big, small):
    return run(big, size=20, bold=True, color="#1f2d3d") + run(NL, raw=True) + run(small, size=9, color="#6b7785")


def title(main, sub):
    return run(main, size=15, bold=True, color="#1f2d3d") + run(NL, raw=True) + run(sub, size=9, color="#6b7785")


FOOT = run("Sources: App Store public reviews, iOS Indonesia, 26 Sep 2026 (names not stored, personal data masked); "
           "Bank Indonesia Tabel 5e. Classifier precision 0.80 / recall 0.64 on 100 held-out hand labels. Independent analysis.",
           size=8, color="#8a94a0")

D1, D2, D3 = "1 Market & Scorecard", "2 Complaint Map", "3 Priorities for DANA"
dash = [
    dashboard(D1, [
        z_text(title("E-money grew 4.3x since 2021, but satisfaction is split in two",
                     "GoPay, DANA and ShopeePay average 3.2-3.7 stars; OVO and LinkAja sit below 1.6, with most reviews at 1 star"),
               615, 600, 78000, 10600),
        z_param(PARAM_LENS, 79000, 600, 20385, 10600),
        *[z_text(kpi(big, small), 615 + i * 24700, 11400, 24400, 12800, bg="#f4f6f8") for i, (big, small) in enumerate([
            ("32.0 bn", "transactions in 2025 (3.9x vs 2021)"),
            ("Rp 3,362 tn", "value in 2025 (4.3x vs 2021)"),
            ("1.07 bn", "e-money instruments (1.9x)"),
            ("Rp 105k", "average per transaction, 2025"),
        ])],
        z_sheet("Average rating per app", 615, 24500, 30000, 35000),
        z_sheet("Star-rating mix per app", 30900, 24500, 36000, 35000),
        z_sheet("Reviews per day", 67200, 24500, 32185, 35000),
        z_sheet("E-money transaction value", 615, 60000, 50000, 33000),
        z_text(run("What this means", size=12, bold=True, color="#1f2d3d") + run(NL, raw=True)
               + run("- The market is growing fast (transactions 3.9x, value 4.3x since 2021), so trust and reliability, "
                     "not awareness, decide who wins.", size=10) + run(NL, raw=True)
               + run("- DANA gets the most reviews per day (20) and a 3.3 average: large, active, but 36% of its latest "
                     "reviews are still 1 star.", size=10) + run(NL, raw=True)
               + run("- Apps are compared as equal-size snapshots of their latest 500 reviews. Those 500 reviews cover "
                     "25 days for DANA but 847 days for LinkAja, so no cross-app time trend is claimed.", size=10),
               51000, 60000, 48385, 33000, bg="#f4f6f8"),
        z_text(FOOT, 615, 93500, 98770, 6000),
    ]),
    dashboard(D2, [
        z_text(title("Failed transfers, poor support and login/OTP issues drive most complaints",
                     "Click any cell in the heatmap to filter the star ratings and the actual reviews on the right. Click again to clear."),
               615, 600, 98770, 10600),
        z_param(PARAM_APP, 615, 11400, 16000, 7000),
        z_param(PARAM_LENS, 16800, 11400, 16000, 7000),
        z_sheet("Review KPIs", 33000, 11400, 66385, 7000, show_title=False),
        z_sheet("Complaint heatmap", 615, 18600, 54000, 74400),
        z_sheet("Star ratings of selection", 55000, 18600, 44385, 21400),
        z_sheet("What users wrote", 55000, 40500, 44385, 52500),
        z_text(FOOT, 615, 93500, 98770, 6000),
    ]),
    dashboard(D3, [
        z_text(title("DANA's standout problem is Paylater: 12.8% of its reviews vs 5.6% at peers",
                     "DANA beats the peer average on failed transfers, customer service and login, but trails on Paylater and lost balance / auto-debit"),
               615, 600, 98770, 10600),
        z_param(PARAM_APP, 615, 11400, 16000, 7000),
        z_param(PARAM_LENS, 16800, 11400, 16000, 7000),
        z_sheet("Priority matrix", 615, 18600, 49000, 45400),
        z_sheet("DANA vs peers", 50000, 11400, 49385, 52600),
        z_sheet("DANA by app version", 615, 64500, 49000, 29000),
        z_text(run("Recommendations for DANA", size=12, bold=True, color="#1f2d3d") + run(NL, raw=True)
               + run("1. DANA Cicil: users see a credit limit they cannot use. Show why and when it unlocks, in the app. "
                     "Biggest gap vs peers (+7.3 pp).", size=9) + run(NL, raw=True)
               + run("2. Auto-debit and bills: confirm before each debit and send a receipt, to cut 'balance lost' "
                     "reviews (+2.0 pp vs peers; 90% of them are 1 star).", size=9) + run(NL, raw=True)
               + run("3. Watch release 2.145.0: rating fell to 3.05 from about 3.4, with more customer-service and fee "
                     "complaints. Not yet significant (p = 0.10); track the next release.", size=9),
               50000, 64500, 49385, 29000, bg="#f4f6f8"),
        z_text(FOOT, 615, 93500, 98770, 6000),
    ]),
]


def action(caption, dashboard_name, source, targets_excluded):
    return f"""<action caption='{q(caption)}' name='[Action_{uuid.uuid4().hex.upper()}]'>
<activation auto-clear='true' type='on-select' />
<source dashboard='{q(dashboard_name)}' type='sheet' worksheet='{q(source)}' />
<command command='tsc:tsl-filter'>
<param name='exclude' value='{q(",".join(targets_excluded))}' />
<param name='special-fields' value='all' />
<param name='target' value='{q(dashboard_name)}' />
</command></action>"""


ACTIONS = [action("Click a heatmap cell to read those reviews", D2, "Complaint heatmap", ["Review KPIs"])]


CARDS = ("<cards><edge name='left'><strip size='160'><card type='pages' /><card type='filters' /><card type='marks' />"
         "</strip></edge><edge name='top'><strip size='2147483647'><card type='columns' /></strip>"
         "<strip size='2147483647'><card type='rows' /></strip><strip size='30'><card type='title' /></strip></edge></cards>")


DASH_SHEETS = {
    D1: ["Average rating per app", "Star-rating mix per app", "Reviews per day", "E-money transaction value"],
    D2: ["Review KPIs", "Complaint heatmap", "Star ratings of selection", "What users wrote"],
    D3: ["Priority matrix", "DANA vs peers", "DANA by app version"],
}


def windows():
    names = [n for n in ("E-money transaction value", "Average rating per app", "Star-rating mix per app",
                         "Reviews per day", "Review KPIs", "Complaint heatmap", "Star ratings of selection",
                         "What users wrote", "Priority matrix", "DANA vs peers", "DANA by app version")]
    w = "".join(f"<window class='worksheet' name='{q(n)}'>{CARDS}<simple-id uuid='{uid()}' /></window>" for n in names)
    for i, (d, sheets) in enumerate(DASH_SHEETS.items()):
        mx = "maximized='true' " if i == 0 else ""
        vp = "".join(f"<viewpoint name='{q(n)}'>" + ("" if n == "What users wrote" else "<zoom type='entire-view' />") + "</viewpoint>" for n in sheets)
        w += f"<window class='dashboard' {mx}name='{q(d)}'><viewpoints>{vp}</viewpoints><active id='-1' /><simple-id uuid='{uid()}' /></window>"
    return f"<windows source-height='30'>{w}</windows>"


def main():
    xml = f"""<?xml version='1.0' encoding='utf-8' ?>
<workbook locale='en_US' original-version='18.1' source-build='2026.2.3 (20262.26.0912.1023)' source-platform='win' version='18.1' xmlns:user='http://www.tableausoftware.com/xml/user'>
<document-format-change-manifest><AccessibleZoneTabOrder /><AnimationOnByDefault /><MarkAnimation /><ObjectModelEncapsulateLegacy /><ObjectModelExtractV2 /><ObjectModelTableType /><SchemaViewerObjectModel /><SetMembershipControl /><SheetIdentifierTracking /><SortTagCleanup /><VConnDownstreamExtractsWithWarnings /><WindowsPersistSimpleIdentifiers /></document-format-change-manifest>
<preferences><preference name='ui.encoding.shelf.height' value='24' /><preference name='ui.shelf.height' value='26' /></preferences>
<datasources>
{parameters()}
{datasource(R, 'dashboard_reviews', 'dashboard_reviews.csv', REVIEW_COLS, REVIEW_CALCS, REVIEW_STYLE, REVIEW_INSTANCES)}
{datasource(M, 'dashboard_market', 'dashboard_market.csv', MARKET_COLS, MARKET_CALCS)}
</datasources>
<actions>{''.join(ACTIONS)}</actions>
<worksheets>{''.join(S.values())}</worksheets>
<dashboards>{''.join(dash)}</dashboards>
{windows()}
</workbook>"""
    minidom.parseString(xml.encode("utf-8"))  # fails loudly on malformed XML
    out = HERE / "ewallet_review_insights.twb"
    out.write_text(xml, encoding="utf-8")
    print(f"wrote {out.name} ({len(xml) / 1024:.0f} KB), {len(S)} sheets, {len(dash)} dashboards")


if __name__ == "__main__":
    main()
