# iPhone の共有シート用ショートカット「家計簿に追加」を作る（署名は macOS の shortcuts sign）
# PayPay の取引詳細の共有ボタン（画像・文字）、スクリーンショット、取引履歴CSV のどれを送っても記録する
# 使い方: python3 build.py <Apps Script の exec URL>   ※ URL 入りのファイルはコミットしない
import plistlib, uuid, subprocess, sys

U = lambda: str(uuid.uuid4()).upper()

def out(uid, name):
    return {'Type': 'ActionOutput', 'OutputUUID': uid, 'OutputName': name}
def att(a):
    return {'Value': a, 'WFSerializationType': 'WFTextTokenAttachment'}
def text(s='', a=None):
    v = {'string': s}
    if a is not None:
        v = {'string': '￼', 'attachmentsByRange': {'{0, 1}': a}}
    return {'Value': v, 'WFSerializationType': 'WFTextTokenString'}
def dic(items):
    return {'Value': {'WFDictionaryFieldValueItems': [{'WFItemType': 0, 'WFKey': text(k), 'WFValue': v} for k, v in items]},
            'WFSerializationType': 'WFDictionaryFieldValue'}
INPUT = {'Type': 'ExtensionInput'}

u0, uImg, uOcr, uTxt, uIf, uPost = U(), U(), U(), U(), U(), U()
g = U()  # もし〜の組
actions = [
    # 0. 送り先（Apps Script の URL）
    {'WFWorkflowActionIdentifier': 'is.workflow.actions.gettext',
     'WFWorkflowActionParameters': {'UUID': u0, 'WFTextActionText': ''}},
    # 1. 共有されたものから画像を取り出す
    {'WFWorkflowActionIdentifier': 'is.workflow.actions.detect.images',
     'WFWorkflowActionParameters': {'UUID': uImg, 'WFInput': att(INPUT)}},
    # 2. 画像があれば文字を読み取る、なければ（文字・CSV）そのまま文字にする
    {'WFWorkflowActionIdentifier': 'is.workflow.actions.conditional',
     'WFWorkflowActionParameters': {'GroupingIdentifier': g, 'WFControlFlowMode': 0, 'WFCondition': 100,
                                    'WFInput': {'Type': 'Variable', 'Variable': att(out(uImg, 'Images'))}}},
    {'WFWorkflowActionIdentifier': 'is.workflow.actions.extracttextfromimage',
     'WFWorkflowActionParameters': {'UUID': uOcr, 'WFImage': att(out(uImg, 'Images'))}},
    {'WFWorkflowActionIdentifier': 'is.workflow.actions.conditional',
     'WFWorkflowActionParameters': {'GroupingIdentifier': g, 'WFControlFlowMode': 1}},
    {'WFWorkflowActionIdentifier': 'is.workflow.actions.detect.text',
     'WFWorkflowActionParameters': {'UUID': uTxt, 'WFInput': att(INPUT)}},
    {'WFWorkflowActionIdentifier': 'is.workflow.actions.conditional',
     'WFWorkflowActionParameters': {'GroupingIdentifier': g, 'WFControlFlowMode': 2, 'UUID': uIf}},
    # 3. Apps Script に送って記録
    {'WFWorkflowActionIdentifier': 'is.workflow.actions.downloadurl',
     'WFWorkflowActionParameters': {'UUID': uPost, 'WFURL': text(a=out(u0, 'Text')), 'WFHTTPMethod': 'POST', 'WFHTTPBodyType': 'JSON', 'ShowHeaders': False,
                                    'WFJSONValues': dic([('action', text('quickAdd')), ('text', text(a=out(uIf, 'If Result'))),
                                                         ('source', text('PayPay')), ('format', text('text'))])}},
    # 4. 結果を表示
    {'WFWorkflowActionIdentifier': 'is.workflow.actions.showresult',
     'WFWorkflowActionParameters': {'Text': text(a=out(uPost, 'Contents of URL'))}},
]
wf = {
    'WFWorkflowActions': actions,
    'WFWorkflowClientVersion': '3218.0.4.100',
    'WFWorkflowMinimumClientVersion': 900,
    'WFWorkflowMinimumClientVersionString': '900',
    'WFWorkflowHasShortcutInputVariables': True,
    'WFWorkflowIcon': {'WFWorkflowIconStartColor': 4292093695, 'WFWorkflowIconGlyphNumber': 59511},
    'WFWorkflowImportQuestions': [{'ActionIndex': 0, 'Category': 'Parameter', 'DefaultValue': '', 'ParameterKey': 'WFTextActionText',
                                   'Text': '家計簿の 設定 → 連携 で使っている Apps Script の URL（https://script.google.com/…/exec）を貼ってください'}],
    'WFWorkflowInputContentItemClasses': ['WFImageContentItem', 'WFGenericFileContentItem', 'WFPDFContentItem', 'WFStringContentItem',
                                          'WFRichTextContentItem', 'WFURLContentItem'],
    'WFWorkflowOutputContentItemClasses': [],
    'WFWorkflowTypes': ['ActionExtension'],
    'WFQuickActionSurfaces': [],
    'WFWorkflowHasOutputFallback': False,
}
if len(sys.argv) > 1:   # URL を埋め込んで作る（追加時の質問なし）
    actions[0]['WFWorkflowActionParameters']['WFTextActionText'] = sys.argv[1]
    wf['WFWorkflowImportQuestions'] = []
with open('unsigned.shortcut', 'wb') as f:
    plistlib.dump(wf, f, fmt=plistlib.FMT_BINARY)
r = subprocess.run(['shortcuts', 'sign', '-m', 'anyone', '-i', 'unsigned.shortcut', '-o', '家計簿に追加.shortcut'], capture_output=True, text=True)
print(r.returncode, r.stdout, r.stderr)
