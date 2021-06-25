# <a name="office-scripts-api-documentation-tools"></a>Officeスクリプト API ドキュメント ツール

これらのツールは、SCripts Officeの背後にあるチームのサポートに役立ちます。 このフォルダーでツールを実行するには、次の手順に従います。

## <a name="coverage-tester"></a>coverage-tester

このツールは、各 API のドキュメント範囲の概要を示します。 各 API は、ドキュメントの品質とサンプル コードの有無について評価されます。 品質指標はまだ開発中です。

このツールの出力はファイル `.csv` です。

### <a name="coverage-tester-instructions"></a>coverage-tester 命令

1. リポジトリを複製またはフォークします。
1. コマンド ウィンドウで、 `/office-scripts-docs-reference/generate-docs/tools`
1. `npm install` を実行します。
1. `npm run build` を実行します。
1. `node coverage-tester` を実行します。
1. Open "API Coverage Report.csv"
