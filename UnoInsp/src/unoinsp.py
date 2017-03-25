#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import re #正規表現モジュール。
import platform  # OS名の取得に使用。
from com.sun.star.uno.TypeClass import SERVICE, INTERFACE, PROPERTY, INTERFACE_METHOD, INTERFACE_ATTRIBUTE
from com.sun.star.beans import PropertyValue
CSS = "com.sun.star"  # IDL名の先頭から省略する部分。
REG_IDL = re.compile(r'(?<!\w)\.[\w\.]+')  # IDL名を抽出する正規表現オブジェクト。
REG_I = re.compile(r'(?<!\w)\.[\w\.]+\.X[\w]+')  # インターフェイス名を抽出する正規表現オブジェクト。
REG_E = re.compile(r'(?<!\w)\.[\w\.]+\.[\w]+Exception')  # 例外名を抽出する正規表現オブジェクト。
ST_OMI = {'.uno.XInterface', '.lang.XTypeProvider', '.lang.XServiceInfo', '.uno.XWeak', '.lang.XComponent', '.lang.XInitialization', '.lang.XMain', '.uno.XAggregation', '.lang.XUnoTunnel'}  # 結果を出力しないインターフェイス名の集合の初期値。
LST_KEY = ["SERVICE", "INTERFACE", "PROPERTY", "INTERFACE_METHOD", "INTERFACE_ATTRIBUTE"]  # dic_fnのキーのリスト。
class ObjInsp:  # XSCRIPTCONTEXTを引数にしてインスタンス化する。第二引数をTrueにするとローカルのAPIリファレンスへのリンクになる。
    def __init__(self, XSCRIPTCONTEXT, offline=False):
        ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストを取得。
        self.st_omi = set()  # 出力を抑制するインターフェイスの集合。
        self.stack = list()  # スタック。
        self.lst_output = list()  # 出力行を収納するリスト。
        self.dic_fn = dict()  # 出力方法を決める関数を入れる辞書。
        self.prefix = "http://api.libreoffice.org/docs/idl/ref/" if not offline else "file://" + get_path(ctx) + "/sdk/docs/idl/ref/"  # offlineがTrueのときはローカルのAPIリファレンスへのリンクを張る。
        self.tdm = ctx.getByName('/singletons/com.sun.star.reflection.theTypeDescriptionManager')  # TypeDescriptionManagerをシングルトンでインスタンス化。
    def tree(self, obj, lst_supr=None):  # 修飾無しでprint()で出力。PyCharmでの使用を想定。
        self._init(lst_supr)  # 初期化関数
        self.dic_fn = dict(zip(LST_KEY, [self.lst_output.append for i in range(len(LST_KEY))]))  # すべてself.lst_output.appendする。
        if isinstance(obj, str): # objが文字列(IDL名)のとき
            self._ext_desc_idl(obj)
        else:  # objが文字列以外の時
            self._ext_desc(obj)
        print("\n".join(self.lst_output))
    def itree(self, obj, lst_supr=None):  # アンカータグをつけて出力。IPython Notebookでの使用を想定
        self._init(lst_supr)  # 初期化関数
        self._output_setting()  # IDL名にリンクをつけて出力するための設定。
        self.lst_output.append("<tt>")  # 等幅フォントのタグを指定。
        if isinstance(obj, str):  # objが文字列(IDL名)のとき
            self._ext_desc_idl(obj)
        else:  # objが文字列以外の時
            self._ext_desc(obj)
        self.lst_output.append("</tt>")  # 等速フォントのタグを閉じる。
        from IPython.display import display, HTML  # IPython Notebook用
        display(HTML("<br/>".join(self.lst_output)))  # IPython Notebookに出力。
    def wtree(self, obj, lst_supr=None):  # ウェブブラウザの新たなタブに出力。マクロやPyCharmでの使用を想定。
        self._init(lst_supr)  # 初期化関数
        self._output_setting()  # IDL名にリンクをつけて出力するための設定。
        self.lst_output.append('<!DOCTYPE html><html><head><meta http-equiv="content-language" content="ja"><meta charset="UTF-8"></head><body><tt>')  # 出力行を収納するリストを初期化。等幅フォントのタグを指定。
        if isinstance(obj, str):  # objが文字列(IDL名)のとき
            self._ext_desc_idl(obj)
        else:
            self._ext_desc(obj)
        self.lst_output = [i + "<br>" for i in self.lst_output]  # リストの各要素の最後に改行タグを追加する。
        self.lst_output.append("</tt></body></html>")  # 等速フォントのタグを閉じる。
        with open('workfile.html', 'w', encoding='UTF-8') as f:  # htmlファイルをUTF-8で作成。すでにあるときは上書き。
            f.writelines(self.lst_output)  # シークエンスデータをファイルに書き出し。
            import webbrowser
            webbrowser.open_new_tab(f.name)  # デフォルトのブラウザの新しいタブでhtmlファイルを開く。
    def _init(self, lst_supr):  # 初期化関数。出力を抑制するインターフェイス名のリストを引数とする。
#         self.st_omi = ST_OMI.copy()  # 結果を出力しないインターフェイス名の集合の初期化。
        self.lst_output = list()  # 出力行を収納するリストを初期化。
        if lst_supr:  # 第2引数があるとき
            if isinstance(lst_supr, list):  # lst_suprがリストのとき
                st_supr = set([i.replace(CSS, "") for i in lst_supr])  # "com.sun.star"を省いて集合に変換。
                if "core" in st_supr:  # coreというキーワードがあるときはST_OMIの要素に置換する。
                    st_supr.remove("core")
                    st_supr.update(ST_OMI)
#                 self.st_omi = st_supr.symmetric_difference(ST_OMI)  # デフォルトでcoreインターフェースを出力しないとき。lst_suprとST_OMIに共通しない要素を取得。
                self.st_omi = st_supr  # デフォルトですべて出力するとき。
            else:  # 引数がリスト以外のとき
                self.lst_output.append(_("第2引数はIDLインターフェイス名のリストで指定してください。"))
        self.stack = list()  # スタックを初期化。
    def _output_setting(self):  # IDL名にリンクをつけて出力するための設定。
        self.dic_fn = dict(zip(LST_KEY, [self._fn for i in range(len(LST_KEY))]))  # 一旦すべての値をself._fnにする。
        self.dic_fn["SERVICE"] = self._fn_s  # サービス名を出力するときの関数を指定。
        self.dic_fn["INTERFACE"] = self._fn_i  # インターフェイス名を出力するときの関数を指定。
    def _fn_s(self, item_with_branch):  # サービス名にアンカータグをつける。
        self._make_link("service", REG_IDL, item_with_branch)
    def _fn_i(self, item_with_branch):  # インターフェイス名にアンカータグをつける。
        self._make_link("interface", REG_I, item_with_branch)
    def _make_link(self, type, regex, item_with_branch):
        idl = regex.findall(item_with_branch)  # 正規表現でIDL名を抽出する。
        if idl:
            lnk = "<a href='" + self.prefix + type + "com_1_1sun_1_1star" + idl[0].replace(".", "_1_1") + ".html' target='_blank'>" + idl[0] + "</a>"  # サービス名のアンカータグを作成。
            self.lst_output.append(item_with_branch.replace(" ", "&nbsp;").replace(idl[0], lnk))  # 半角スペースを置換後にサービス名をアンカータグに置換。
        else:
            self.lst_output.append(item_with_branch.replace(" ", "&nbsp;"))  # 半角スペースを置換。
    def _fn(self, item_with_branch):  # サービス名とインターフェイスを出力するときの関数。
        idl = set(REG_IDL.findall(item_with_branch)) # 正規表現でIDL名を抽出する。
        inf = REG_I.findall(item_with_branch) # 正規表現でインターフェイス名を抽出する。
        exc = REG_E.findall(item_with_branch) # 正規表現で例外名を抽出する。
        idl.difference_update(inf, exc)  # IDL名のうちインターフェイス名と例外名を除く。
        idl = list(idl)  # 残ったIDL名はすべてStructと考えて処理する。
        item_with_branch = item_with_branch.replace(" ", "&nbsp;")  # まず半角スペースをHTMLに置換する。
        for i in inf:  # インターフェイス名があるとき。
            item_with_branch = self._make_anchor("interface", i, item_with_branch)
        for i in exc:  # 例外名があるとき。
            item_with_branch = self._make_anchor("exception", i, item_with_branch)
        for i in idl:  # インターフェイス名と例外名以外について。
            item_with_branch = self._make_anchor("struct", i, item_with_branch)
        self.lst_output.append(item_with_branch)
    def _make_anchor(self, type, i, item_with_branch):
        lnk = "<a href='" + self.prefix + type + "com_1_1sun_1_1star" + i.replace(".", "_1_1") + ".html' target='_blank'　style='text-decoration:none;'>" + i + "</a>"  # 下線はつけない。
        return item_with_branch.replace(i, lnk)
    def _ext_desc(self, obj, flag=False):  #  オブジェクトがサポートするIDLから末裔を抽出する。flagはオブジェクトが直接インターフェイスをもっているときにTrueになるフラグ。
        self.lst_output.append("pyuno object")  # treeの根に表示させるもの。
        if hasattr(obj, "getSupportedServiceNames"):  # オブジェクトがサービスを持っているとき。
            flag = True if hasattr(obj, "getTypes") else False  # サービスを介さないインターフェイスがあるときフラグを立てる。
            st_ss = set([i for i in obj.getSupportedServiceNames() if self._idl_check(i)])  # サポートサービス名一覧からTypeDescriptionオブジェクトを取得できないサービス名を除いた集合を得る。
            st_sups = set()  # 親サービスを入れる集合。
            if len(st_ss) > 1:  # サポートしているサービス名が複数ある場合。
                self.stack = [self.tdm.getByHierarchicalName(i) for i in st_ss]  # サポートサービスのTypeDescriptionオブジェクトをスタックに取得。
                while self.stack:  # スタックがある間実行。
                    j = self.stack.pop()  # サービスのTypeDescriptionオブジェクトを取得。
                    t_std = j.getMandatoryServices() + j.getOptionalServices()  # 親サービスのタプルを取得。
                    lst_std = [i for i in t_std if not i.Name in st_sups]  # 親サービスのTypeDescriptionオブジェクトのうち既に取得した親サービスにないものだけを取得。
                    self.stack.extend(lst_std)  # スタックに新たなサービスのTypeDescriptionオブジェクトのみ追加。
                    st_sups.update([i.Name for i in lst_std])  # 既に取得した親サービス名の集合型に新たに取得したサービス名を追加。
            st_ss.difference_update(st_sups)  # オブジェクトのサポートサービスのうち親サービスにないものだけにする=これがサービスの末裔。
            self.stack = [self.tdm.getByHierarchicalName(i) for i in st_ss]  # TypeDescriptionオブジェクトに変換。
            if self.stack: self.stack.sort(key=lambda x: x.Name, reverse=True)  # Name属性で降順に並べる。
            self._make_tree(flag)
        if hasattr(obj, "getTypes"):  # サポートしているインターフェイスがある場合。
            flag = False
            st_si = set([i.typeName.replace(CSS, "") for i in obj.getTypes()])  # サポートインターフェイス名を集合型で取得。
            lst_si = sorted(list(st_si.difference(self.st_omi)), reverse=True)  # 除外するインターフェイス名を除いて降順のリストにする。
            self.stack = [self.tdm.getByHierarchicalName(i if not i[0] == "." else CSS + i) for i in lst_si]  # TypeDescriptionオブジェクトに変換。CSSが必要。
            self._make_tree(flag)
        if not (hasattr(obj, "getSupportedServiceNames") or hasattr(obj, "getTypes")):  # サポートするサービスやインターフェイスがないとき。
            self.lst_output.append(_("サポートするサービスやインターフェイスがありません。"))
    def _ext_desc_idl(self, idl):  # objがIDL名のとき。
        if idl[0] == ".":  # 先頭が.で始まっているとき
            idl = CSS + idl  # com.sun.starが省略されていると考えて、com.sun.starを追加する。
        j = self._idl_check(idl)  # IDL名からTypeDescriptionオブジェクトを取得。
        if j:
            typcls = j.getTypeClass()  # jのタイプクラスを取得。
            if typcls == INTERFACE or typcls == SERVICE:  # jがサービスかインターフェイスのとき。
                self.lst_output.append(idl)  # treeの根にIDL名を表示
                self.stack = [j]  # TypeDescriptionオブジェクトをスタックに取得
                self._make_tree(flag=False)
            else:  # サービスかインターフェイス以外のときは未対応。
                self.lst_output.append(idl + _("はサービス名またはインターフェイス名ではないので未対応です。"))
        else:  # TypeDescriptionオブジェクトを取得できなかったとき。
            self.lst_output.append(idl + _("はIDL名ではありません。"))
    def _idl_check(self, idl): # IDL名からTypeDescriptionオブジェクトを取得。
            try:
                j = self.tdm.getByHierarchicalName(idl)  # IDL名からTypeDescriptionオブジェクトを取得。
            except:
                j = None
            return j
    def _make_tree(self, flag):  # 末裔から祖先を得て木を出力する。flagはオブジェクトが直接インターフェイスをもっているときにTrueになるフラグ。
        if self.stack:  # 起点となるサービスかインターフェイスがあるとき。
            lst_level = [1 for i in self.stack]  # self.stackの要素すべてについて階層を取得。
            indent = "      "  # インデントを設定。
            m = 0  # 最大文字数を初期化。
            inout_dic = {(True, False): "[in]", (False, True): "[out]", (True, True): "[inout]"}  # メソッドの引数のinout変換辞書。
            t_itd = tuple()  # インターフェイスのTypeDescriptionオブジェクトの入れ物を初期化。
            t_md = tuple()  # メソッドのTypeDescriptionオブジェクトの入れ物を初期化。
            t_spd = tuple()  # サービス属性のTypeDescriptionオブジェクトの入れ物を初期化。
            while self.stack:  # スタックがある間実行。
                j = self.stack.pop()  # スタックからTypeDescriptionオブジェクトをpop。
                level = lst_level.pop()  # jの階層を取得。
                typcls = j.getTypeClass()  # jのタイプクラスを取得。
                branch = ["", ""]  # 枝をリセット。jがサービスまたはインターフェイスのときjに直接つながる枝は1番の要素に入れる。それより左の枝は0番の要素に加える。
                if level > 1:  # 階層が2以上のとき。
                    p = 1  # 処理開始する階層を設定。
                    if flag:  # サービスを介さないインターフェイスがあるとき
                        branch[0] = "│   "  # 階層1に立枝をつける。
                        p = 2  # 階層2から処理する。
                    for i in range(p, level):  # 階層iから出た枝が次行の階層i-1の枝になる。
                        branch[0] += "│   " if i in lst_level else indent  # iは枝の階層ではなく、枝のより上の行にあるその枝がでた階層になる。
                if typcls == INTERFACE or typcls == SERVICE:  # jがサービスかインターフェイスのとき。
                    if level == 1 and flag:  # 階層1かつサービスを介さないインターフェイスがあるとき
                        branch[1] = "├─"  # 階層1のときは下につづく分岐をつける。
                    else:
                        branch[1] = "├─" if level in lst_level else "└─"  # スタックに同じ階層があるときは"├─" 。
                else:  # jがインターフェイスかサービス以外のとき。
                    branch[1] = indent  # 横枝は出さない。
                    if level in lst_level:  # スタックに同じ階層があるとき。
                        typcls2 = self.stack[lst_level.index(level)].getTypeClass()  # スタックにある同じ階層のものの先頭の要素のTypeClassを取得。
                        if typcls2 == INTERFACE or typcls2 == SERVICE: branch[1] = "│   "  # サービスかインターフェイスのとき。横枝だったのを縦枝に書き換える。
                if typcls == INTERFACE_METHOD:  # jがメソッドのとき。
                    typ = j.ReturnType.Name.replace(CSS, "")  # 戻り値の型を取得。
                    if typ[1] == "]": typ = typ.replace("]", "") + "]"  # 属性がシークエンスのとき[]の表記を修正。
                    stack2 = list(j.Parameters)[::-1]  # メソッドの引数について逆順(降順ではない)にスタック2に取得。
                    if not stack2:  # 引数がないとき。
                        branch.append(typ.rjust(m) + "  " + j.MemberName.replace(CSS, "") + "()")  # 「戻り値の型(固定幅mで右寄せ) メソッド名()」をbranchの3番の要素に取得。
                        self.dic_fn["INTERFACE_METHOD"]("".join(branch))  # 枝をつけてメソッドを出力。
                    else:  # 引数があるとき。
                        m3 = max([len(i.Type.Name.replace(CSS, "")) for i in stack2])  # 引数の型の最大文字数を取得。
                        k = stack2.pop()  # 先頭の引数を取得。
                        inout = inout_dic[(k.isIn(), k.isOut())]  # 引数の[in]の判定、[out]の判定
                        typ2 = k.Type.Name.replace(CSS, "")  # 引数の型を取得。
                        if typ2[1] == "]": typ2 = typ2.replace("]", "") + "]"  # 引数の型がシークエンスのとき[]の表記を修正。
                        branch.append(typ.rjust(m) + "  " + j.MemberName.replace(CSS, "") + "( " + inout + " " + typ2.rjust(m3) + " " + k.Name.replace(CSS, ""))  # 「戻り値の型(固定幅で右寄せ)  メソッド名(inout判定　引数の型(固定幅m3で左寄せ) 引数名」をbranchの3番の要素に取得。
                        m2 = len(typ.rjust(m) + "  " + j.MemberName.replace(CSS, "") + "( ")  # メソッドの引数の部分をインデントする文字数を取得。
                        if stack2:  # 引数が複数あるとき。
                            branch.append(",")  # branchの4番の要素に「,」を取得。
                            self.dic_fn["INTERFACE_METHOD"]("".join(branch))  # 枝をつけてメソッド名とその0番の引数を出力。
                            del branch[2:]  # branchの2番以上の要素は破棄する。
                            while stack2:  # 1番以降の引数があるとき。
                                k = stack2.pop()
                                inout = inout_dic[(k.isIn(), k.isOut())]  # 引数の[in]の判定、[out]の判定
                                typ2 = k.Type.Name.replace(CSS, "")  # 引数の型を取得。
                                if typ2[1] == "]": typ2 = typ2.replace("]", "") + "]"  # 引数の型がシークエンスのとき[]の表記を修正。
                                branch.append(" ".rjust(m2) + inout + " " + typ2.rjust(m3) + " " + k.Name.replace(CSS, ""))  # 「戻り値の型とメソッド名の固定幅m2 引数の型(固定幅m3で左寄せ) 引数名」をbranchの2番の要素に取得。
                                if stack2:  # 最後の引数でないとき。
                                    branch.append(",")  # branchの3番の要素に「,」を取得。
                                    self.dic_fn["INTERFACE_METHOD"]("".join(branch))  # 枝をつけて引数を出力。
                                    del branch[2:]  # branchの2番以上の要素は破棄する。
                        t_ex = j.Exceptions  # 例外を取得。
                        if t_ex:  # 例外があるとき。
                            self.dic_fn["INTERFACE_METHOD"]("".join(branch))  # 最後の引数を出力。
                            del branch[2:]  # branchの2番以降の要素を削除。
                            n = ") raises ( "  # 例外があるときに表示する文字列。
                            m4 = len(n)  # nの文字数。
                            stack2 = list(t_ex)  # 例外のタプルをリストに変換。
                            branch.append(" ".rjust(m2 - m4))  # 戻り値の型とメソッド名の固定幅m2からnの文字数を引いて2番の要素に取得。
                            branch.append(n)  # nを3番要素に取得。他の要素と別にしておかないとn in branchがTrueにならない。
                            while stack2:  # stack2があるとき
                                k = stack2.pop()  # 例外を取得。
                                branch.append(k.Name.replace(CSS, ""))  # branchの3番(初回以外のループでは4番)の要素に例外名を取得。
                                if stack2:  # まだ次の例外があるとき。
                                    self.dic_fn["INTERFACE_METHOD"]("".join(branch) + ",")  # 「,」をつけて出力。
                                else:  # 最後の要素のとき。
                                    self.dic_fn["INTERFACE_METHOD"]("".join(branch) + ")")  # 閉じ括弧をつけて出力。
                                if n in branch:  # nが枝にあるときはnのある3番以上のbranchの要素を削除。
                                    del branch[3:]
                                    branch.append(" ".rjust(m4))  # nを削った分の固定幅をbranchの3番の要素に追加。
                                else:
                                    del branch[4:]  # branchの4番以上の要素を削除。
                        else:  # 例外がないとき。
                            self.dic_fn["INTERFACE_METHOD"]("".join(branch) + ")")  # 閉じ括弧をつけて最後の引数を出力。
                else:  # jがメソッド以外のとき。
                    if typcls == INTERFACE:  # インターフェイスのとき。XInterfaceTypeDescription2インターフェイスをもつTypeDescriptionオブジェクト。
                        branch.append(j.Name.replace(CSS, ""))  # インターフェイス名をbranchの2番要素に追加。
                        t_itd = j.getBaseTypes() + j.getOptionalBaseTypes()  # 親インターフェイスを取得。
                        t_md = j.getMembers()  # インターフェイス属性とメソッドのTypeDescriptionオブジェクトを取得。
                        self.dic_fn["INTERFACE"]("".join(branch))  # 枝をつけて出力。
                    elif typcls == PROPERTY:  # サービス属性のとき。
                        typ = j.getPropertyTypeDescription().Name.replace(CSS, "")  # 属性の型
                        if typ[1] == "]": typ = typ.replace("]", "") + "]"  # 属性がシークエンスのとき[]の表記を修正。
                        branch.append(typ.rjust(m) + "  " + j.Name.replace(CSS, ""))  # 型は最大文字数で右寄せにする。
                        self.dic_fn["PROPERTY"]("".join(branch))  # 枝をつけて出力。
                    elif typcls == INTERFACE_ATTRIBUTE:  # インターフェイス属性のとき。
                        typ = j.Type.Name.replace(CSS, "")  # 戻り値の型
                        if typ[1] == "]": typ = typ.replace("]", "") + "]"  # 属性がシークエンスのとき[]の表記を修正。
                        branch.append(typ.rjust(m) + "  " + j.MemberName.replace(CSS, ""))  # 型は最大文字数で右寄せにする。
                        self.dic_fn["INTERFACE_METHOD"]("".join(branch))  # 枝をつけて出力。
                    elif typcls == SERVICE:  # jがサービスのときtdはXServiceTypeDescriptionインターフェイスをもつ。
                        branch.append(j.Name.replace(CSS, ""))  # サービス名をbranchの2番要素に追加。
                        t_std = j.getMandatoryServices() + j.getOptionalServices()  # 親サービスを取得。
                        self.stack.extend(sorted(list(t_std), key=lambda x: x.Name, reverse=True))  # 親サービス名で降順に並べてサービスのTypeDescriptionオブジェクトをスタックに追加。
                        lst_level.extend([level + 1 for i in t_std])  # 階層を取得。
                        itd = j.getInterface()  # new-styleサービスのインターフェイスを取得。TypeDescriptionオブジェクト。
                        if itd:  # new-styleサービスのインターフェイスがあるとき。
                            t_itd = itd,  # XInterfaceTypeDescription2インターフェイスをもつTypeDescriptionオブジェクト。
                        else:  # new-styleサービスのインターフェイスがないときはold-styleサービスのインターフェイスを取得。
                            t_itd = j.getMandatoryInterfaces() + j.getOptionalInterfaces()  # XInterfaceTypeDescriptionインターフェイスをもつTypeDescriptionオブジェクト。
                        t_spd = j.Properties  # サービスからXPropertyTypeDescriptionインターフェイスをもつオブジェクトのタプルを取得。
                        self.dic_fn["SERVICE"]("".join(branch))  # 枝をつけて出力。
                if t_itd:  # 親インターフェイスがあるとき。(TypeDescriptionオブジェクト)
                    lst_itd = [i for i in t_itd if not i.Name.replace(CSS, "") in self.st_omi]  # st_omiを除く。
                    self.stack.extend(sorted(lst_itd, key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
                    lst_level.extend([level + 1 for i in lst_itd])  # 階層を取得。
                    self.st_omi.update([i.Name.replace(CSS, "") for i in lst_itd])  # すでにでてきたインターフェイス名をst_omiに追加して次は使わないようにする。
                    t_itd = tuple()  # インターフェイスのTypeDescriptionオブジェクトの入れ物を初期化。
                if t_md:  # インターフェイス属性とメソッドがあるとき。
                    self.stack.extend(sorted(t_md, key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
                    lst_level.extend([level + 1 for i in t_md])  # 階層を取得。
                    m = max([len(i.ReturnType.Name.replace(CSS, "")) for i in t_md if i.getTypeClass() == INTERFACE_METHOD] + [len(i.Type.Name.replace(CSS, "")) for i in t_md if i.getTypeClass() == INTERFACE_ATTRIBUTE])  # インターフェイス属性とメソッドの型のうち最大文字数を取得。
                    t_md = tuple()  # メソッドのTypeDescriptionオブジェクトの入れ物を初期化。
                if t_spd:  # サービス属性があるとき。
                    self.stack.extend(sorted(list(t_spd), key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
                    lst_level.extend([level + 1 for i in t_spd])  # 階層を取得。
                    m = max([len(i.getPropertyTypeDescription().Name.replace(CSS, "")) for i in t_spd])  # サービス属性の型のうち最大文字数を取得。
                    t_spd = tuple()  # サービス属性のTypeDescriptionオブジェクトの入れ物を初期化。
def get_path(ctx):  # LibreOfficeのインストールパスを得る。
    cp = ctx.getServiceManager().createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
    node = PropertyValue()
    node.Name = 'nodepath'
    node.Value = 'org.openoffice.Setup/Product'  # share/registry/main.xcd内のノードパス。
    ca = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
    o_name = ca.getPropertyValue('ooName')  # ooNameプロパティからLibreOfficeの名前を得る。
    o_setupversion = ca.getPropertyValue('ooSetupVersion')  # ooSetupVersionからLibreOfficeのメジャーバージョンとマイナーバージョンの番号を得る。
    os = platform.system()  # OS名を得る。
    path = "/opt/" if os == "Linux" else ""  # LinuxのときのLibreOfficeのインストール先フォルダ。
    return path + o_name.lower() + o_setupversion