# unity-project-guide
untiyの講義資料を管理するリポジトリです。

## OfficeファイルのGit管理

`.pptx`, `.docx`, `.xlsx` などのOffice Open XML形式はzipファイルなので、
Git登録時に `scripts/ooxml_filter.py` で一度展開し、XMLとzipを正規化してから再圧縮します。
これにより、zip内ファイルの順序・タイムスタンプ・XML属性順・Officeが更新しやすい作成者や更新日時などの差分を抑えます。

初回だけ、リポジトリ内で次を実行してください。

```powershell
python scripts/ooxml_filter.py install
```

手動で試す場合は次のように使えます。

```powershell
python scripts/ooxml_filter.py unpack sample.pptx sample-pptx
python scripts/ooxml_filter.py pack sample-pptx sample-normalized.pptx
python scripts/ooxml_filter.py textconv sample.pptx
```
