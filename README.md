# mongodb-certificate-checker
MongoDB certificate verification and scoring project

学生が提出した修了証(PDF)を検証し、Self Directed Learning(SDL)のポイントを
計算してMongoDBに保存、Excelレポートを出力するプログラム

## 必要な環境
- Python 3.7以上
- MongoDB(ローカルで起動していること、localhost:27017)

## セットアップ
\`\`\`
pip install -r requirements.txt
\`\`\`

## 使い方
1. \`python certificate_checker.py\` を実行
2. ウィンドウで、Certificate名・フォルダのパス・Target Completion Date(DD/MM/YYYY)を入力
3. 「実行」を押すと、MongoDBの internship_db.SD_Marks に保存され、
   SD_Marks_Report.xlsx が出力される

## 採点ルール
- 検証OK、期限内(Target Completion Date以内)の提出:2点
- 検証OK、期限後の提出:1点
- 検証NG(名前不一致、Certificate名不一致、IDなしなど):0点