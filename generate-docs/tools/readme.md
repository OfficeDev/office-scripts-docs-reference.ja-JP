# <a name="office-scripts-api-documentation-tools"></a><span data-ttu-id="bfa7e-101">Officeスクリプト API ドキュメント ツール</span><span class="sxs-lookup"><span data-stu-id="bfa7e-101">Office Scripts API Documentation Tools</span></span>

<span data-ttu-id="bfa7e-102">これらのツールは、SCripts Officeの背後にあるチームのサポートに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="bfa7e-102">These tools help support the Office SCripts documentation and the team behind it.</span></span> <span data-ttu-id="bfa7e-103">このフォルダーでツールを実行するには、次の手順に従います。</span><span class="sxs-lookup"><span data-stu-id="bfa7e-103">Follow these instructions to run the tools in this folder.</span></span>

## <a name="coverage-tester"></a><span data-ttu-id="bfa7e-104">coverage-tester</span><span class="sxs-lookup"><span data-stu-id="bfa7e-104">coverage-tester</span></span>

<span data-ttu-id="bfa7e-105">このツールは、各 API のドキュメント範囲の概要を示します。</span><span class="sxs-lookup"><span data-stu-id="bfa7e-105">This tool gives an overview of the documentation coverage for each API.</span></span> <span data-ttu-id="bfa7e-106">各 API は、ドキュメントの品質とサンプル コードの有無について評価されます。</span><span class="sxs-lookup"><span data-stu-id="bfa7e-106">Each API is assessed for documentation quality and the presence of sample code.</span></span> <span data-ttu-id="bfa7e-107">品質指標はまだ開発中です。</span><span class="sxs-lookup"><span data-stu-id="bfa7e-107">The quality metrics are still in development.</span></span>

<span data-ttu-id="bfa7e-108">このツールの出力はファイル `.csv` です。</span><span class="sxs-lookup"><span data-stu-id="bfa7e-108">The output of this tool is a `.csv` file.</span></span>

### <a name="coverage-tester-instructions"></a><span data-ttu-id="bfa7e-109">coverage-tester 命令</span><span class="sxs-lookup"><span data-stu-id="bfa7e-109">coverage-tester Instructions</span></span>

1. <span data-ttu-id="bfa7e-110">リポジトリを複製またはフォークします。</span><span class="sxs-lookup"><span data-stu-id="bfa7e-110">Clone or fork the repo.</span></span>
1. <span data-ttu-id="bfa7e-111">コマンド ウィンドウで、 `/office-scripts-docs-reference/generate-docs/tools`</span><span class="sxs-lookup"><span data-stu-id="bfa7e-111">In a command window, go to `/office-scripts-docs-reference/generate-docs/tools`</span></span>
1. <span data-ttu-id="bfa7e-112">`npm install` を実行します。</span><span class="sxs-lookup"><span data-stu-id="bfa7e-112">Run `npm install`</span></span>
1. <span data-ttu-id="bfa7e-113">`npm run build` を実行します。</span><span class="sxs-lookup"><span data-stu-id="bfa7e-113">Run `npm run build`</span></span>
1. <span data-ttu-id="bfa7e-114">`node coverage-tester` を実行します。</span><span class="sxs-lookup"><span data-stu-id="bfa7e-114">Run `node coverage-tester`</span></span>
1. <span data-ttu-id="bfa7e-115">Open "API Coverage Report.csv"</span><span class="sxs-lookup"><span data-stu-id="bfa7e-115">Open “API Coverage Report.csv”</span></span>
