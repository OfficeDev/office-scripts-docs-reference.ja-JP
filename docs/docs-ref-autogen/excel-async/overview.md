---
title: Office スクリプト API リファレンス
description: Office スクリプトの非同期 JavaScript Api の概要。
ms.date: 06/17/2020
ms.openlocfilehash: b9fa59764b7cb54567adbe05b9671bae9b232972
ms.sourcegitcommit: 163b26a43411ad7f13a01237efe9b8d6de656b47
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44887086"
---
# <a name="office-scripts-async-api-reference"></a><span data-ttu-id="ae05c-103">Office スクリプトの非同期 API リファレンス</span><span class="sxs-lookup"><span data-stu-id="ae05c-103">Office Scripts Async API reference</span></span>

<span data-ttu-id="ae05c-104">Office Scripts Async API は、Office スクリプトのプレビュー段階で作成された古いスクリプトをサポートしています。</span><span class="sxs-lookup"><span data-stu-id="ae05c-104">The Office Scripts Async API supports older scripts made during the Office Scripts preview phase.</span></span> <span data-ttu-id="ae05c-105">これらの以前のスクリプトで使用されていたクラス、メソッド、およびその他の種類の詳細については、このリファレンスドキュメントを使用してください。</span><span class="sxs-lookup"><span data-stu-id="ae05c-105">Use this reference documentation to learn more about the classes, methods, and other types used by these older scripts.</span></span> <span data-ttu-id="ae05c-106">Office スクリプトを使用してアクセスできるすべてのオブジェクトについては、ページの左側にある目次を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ae05c-106">All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ae05c-107">標準の Office スクリプト Api を使用して、新しいスクリプトを作成することを強くお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ae05c-107">We strongly recommend creating new scripts with the standard Office Scripts APIs.</span></span> <span data-ttu-id="ae05c-108">新しいスクリプトを作成または編集している場合は、[非非同期バージョン](?view=office-scripts)の api に切り替えてください。</span><span class="sxs-lookup"><span data-stu-id="ae05c-108">If you're making or editing a new script, please switch to the [non-async version](?view=office-scripts) of the APIs.</span></span>

## <a name="common-classes"></a><span data-ttu-id="ae05c-109">一般的なクラス</span><span class="sxs-lookup"><span data-stu-id="ae05c-109">Common classes</span></span>

<span data-ttu-id="ae05c-110">次の一覧は、Office Scripts オブジェクトモデルの基本事項を示しています。</span><span class="sxs-lookup"><span data-stu-id="ae05c-110">The following list breaks down the basics of the Office Scripts object model.</span></span> <span data-ttu-id="ae05c-111">これにより、共通のクラスと、それらが相互にどのように関係しているかがわかります。</span><span class="sxs-lookup"><span data-stu-id="ae05c-111">This shows the common classes and how they relate to one another.</span></span>

- <span data-ttu-id="ae05c-112">[ブック](/javascript/api/office-scripts/excel/excelscript.workbook)には、1つまたは複数の[ワークシート](/javascript/api/office-scripts/excel/excelscript.worksheet)が含まれて[います。](/javascript/api/office-scripts/excel/excelscript.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="ae05c-112">A [Workbook](/javascript/api/office-scripts/excel/excelscript.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excel/excelscript.worksheet) in a [WorksheetCollection](/javascript/api/office-scripts/excel/excelscript.worksheetcollection).</span></span>
- <span data-ttu-id="ae05c-113">[ワークシート](/javascript/api/office-scripts/excel/excelscript.worksheet) では、[Range](/javascript/api/office-scripts/excel/excelscript.range) オブジェクトを介してセルにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="ae05c-113">A [Worksheet](/javascript/api/office-scripts/excel/excelscript.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excel/excelscript.range) objects.</span></span>
- <span data-ttu-id="ae05c-114">[Range](/javascript/api/office-scripts/excel/excelscript.range) は、連続したセルのグループを表します。</span><span class="sxs-lookup"><span data-stu-id="ae05c-114">A [Range](/javascript/api/office-scripts/excel/excelscript.range) represents a group of contiguous cells.</span></span>
- <span data-ttu-id="ae05c-115">[Range](/javascript/api/office-scripts/excel/excelscript.range) は、[表](/javascript/api/office-scripts/excel/excelscript.table)、[グラフ](/javascript/api/office-scripts/excel/excelscript.chart)、[図形](/javascript/api/office-scripts/excel/excelscript.shape)、およびその他のデータ可視化や組織オブジェクトを作成して配置するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="ae05c-115">[Ranges](/javascript/api/office-scripts/excel/excelscript.range) are used to create and place [Tables](/javascript/api/office-scripts/excel/excelscript.table), [Charts](/javascript/api/office-scripts/excel/excelscript.chart), [Shapes](/javascript/api/office-scripts/excel/excelscript.shape), and other data visualization or organization objects.</span></span>
- <span data-ttu-id="ae05c-116">[ワークシート](/javascript/api/office-scripts/excel/excelscript.worksheet)には、個々のシートにあるデータオブジェクト ( [chartcollection](/javascript/api/office-scripts/excel/excelscript.chartcollection)など) のコレクションが含まれています。</span><span class="sxs-lookup"><span data-stu-id="ae05c-116">A [Worksheet](/javascript/api/office-scripts/excel/excelscript.worksheet) contains collections of those data objects (such as a [ChartCollection](/javascript/api/office-scripts/excel/excelscript.chartcollection)) that are present in the individual sheet.</span></span>
- <span data-ttu-id="ae05c-117">[ブック全体](/javascript/api/office-scripts/excel/excelscript.workbook)の一部のデータオブジェクト ( [tablecollection](/javascript/api/office-scripts/excel/excelscript.tablecollection)など) のコレクションが含まれ[ています](/javascript/api/office-scripts/excel/excelscript.workbook)。</span><span class="sxs-lookup"><span data-stu-id="ae05c-117">[Workbooks](/javascript/api/office-scripts/excel/excelscript.workbook) contain collections of some of those data objects (such as a [TableCollection](/javascript/api/office-scripts/excel/excelscript.tablecollection)) for the entire [Workbook](/javascript/api/office-scripts/excel/excelscript.workbook).</span></span>

<span data-ttu-id="ae05c-118">Office Scripts オブジェクトモデルの詳細については、「 [Office スクリプトの基礎知識](/office/dev/scripts/develop/scripting-fundamentals)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ae05c-118">For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)</span></span>

## <a name="see-also"></a><span data-ttu-id="ae05c-119">関連項目</span><span class="sxs-lookup"><span data-stu-id="ae05c-119">See also</span></span>

- [<span data-ttu-id="ae05c-120">従来のスクリプトをサポートするための Office スクリプト非同期 Api の使用</span><span class="sxs-lookup"><span data-stu-id="ae05c-120">Using the Office Scripts Async APIs to support legacy scripts</span></span>](/office/dev/scripts/develop/excel-async-model)
- [<span data-ttu-id="ae05c-121">Office スクリプトについて</span><span class="sxs-lookup"><span data-stu-id="ae05c-121">About Office Scripts</span></span>](/office/dev/scripts/overview/excel)
- [<span data-ttu-id="ae05c-122">Excel on the web で Office スクリプトを記録、編集、作成する</span><span class="sxs-lookup"><span data-stu-id="ae05c-122">Record, edit, and create Office Scripts in Excel on the web</span></span>](/office/dev/scripts/tutorials/excel-tutorial)
- [<span data-ttu-id="ae05c-123">Excel on the web での Office スクリプトのスクリプトの基本事項</span><span class="sxs-lookup"><span data-stu-id="ae05c-123">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](/office/dev/scripts/develop/scripting-fundamentals)
