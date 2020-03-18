---
title: Office スクリプト API リファレンス
description: Office スクリプト JavaScript Api の概要
ms.date: 12/12/2019
ms.openlocfilehash: b3e75cbecac13f41c83f019c681b0682571ead41
ms.sourcegitcommit: 2834ee43a14a4b0953d5fb22749c115a42960e20
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2020
ms.locfileid: "42704476"
---
# <a name="office-scripts-api-reference"></a><span data-ttu-id="c0d0d-103">Office スクリプト API リファレンス</span><span class="sxs-lookup"><span data-stu-id="c0d0d-103">Office Scripts API reference</span></span>

<span data-ttu-id="c0d0d-104">Office スクリプト API を使用すると、web 上の Excel で一般的なタスクを自動化できます。</span><span class="sxs-lookup"><span data-stu-id="c0d0d-104">The Office Scripts API lets you automate common tasks in Excel on the web.</span></span> <span data-ttu-id="c0d0d-105">このリファレンスドキュメントを使用して、スクリプトで使用できるクラス、メソッド、およびその他の種類の詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="c0d0d-105">Use this reference documentation to learn more about the classes, methods, and other types available for your scripts.</span></span> <span data-ttu-id="c0d0d-106">Office スクリプトを使用してアクセスできるすべてのオブジェクトについては、ページの左側にある目次を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c0d0d-106">All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.</span></span>

## <a name="common-classes"></a><span data-ttu-id="c0d0d-107">一般的なクラス</span><span class="sxs-lookup"><span data-stu-id="c0d0d-107">Common classes</span></span>

<span data-ttu-id="c0d0d-108">次の一覧は、Office Scripts オブジェクトモデルの基本事項を示しています。</span><span class="sxs-lookup"><span data-stu-id="c0d0d-108">The following list breaks down the basics of the Office Scripts object model.</span></span> <span data-ttu-id="c0d0d-109">これにより、共通のクラスと、それらが相互にどのように関係しているかがわかります。</span><span class="sxs-lookup"><span data-stu-id="c0d0d-109">This shows the common classes and how they relate to one another.</span></span>

- <span data-ttu-id="c0d0d-110">[ブック](/javascript/api/office-scripts/excel/excel.workbook)には、1つまたは複数の[ワークシート](/javascript/api/office-scripts/excel/excel.worksheet)が含まれて[います。](/javascript/api/office-scripts/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="c0d0d-110">A [Workbook](/javascript/api/office-scripts/excel/excel.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excel/excel.worksheet) in a [WorksheetCollection](/javascript/api/office-scripts/excel/excel.worksheetcollection).</span></span>
- <span data-ttu-id="c0d0d-111">[ワークシート](/javascript/api/office-scripts/excel/excel.worksheet)は、 [Range](/javascript/api/office-scripts/excel/excel.range)オブジェクトを通じてセルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c0d0d-111">A [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excel/excel.range) objects.</span></span>
- <span data-ttu-id="c0d0d-112">[範囲](/javascript/api/office-scripts/excel/excel.range)は、連続したセルのグループを表します。</span><span class="sxs-lookup"><span data-stu-id="c0d0d-112">A [Range](/javascript/api/office-scripts/excel/excel.range) represents a group of contiguous cells.</span></span>
- <span data-ttu-id="c0d0d-113">[範囲](/javascript/api/office-scripts/excel/excel.range)は、[表](/javascript/api/office-scripts/excel/excel.table)、[グラフ](/javascript/api/office-scripts/excel/excel.chart)、[図形](/javascript/api/office-scripts/excel/excel.shape)、およびその他のデータビジュアライゼーションまたは組織オブジェクトを作成して配置するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="c0d0d-113">[Ranges](/javascript/api/office-scripts/excel/excel.range) are used to create and place [Tables](/javascript/api/office-scripts/excel/excel.table), [Charts](/javascript/api/office-scripts/excel/excel.chart), [Shapes](/javascript/api/office-scripts/excel/excel.shape), and other data visualization or organization objects.</span></span>
- <span data-ttu-id="c0d0d-114">[ワークシート](/javascript/api/office-scripts/excel/excel.worksheet)には、個々のシートにあるデータオブジェクト ( [chartcollection](/javascript/api/office-scripts/excel/excel.chartcollection)など) のコレクションが含まれています。</span><span class="sxs-lookup"><span data-stu-id="c0d0d-114">A [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet) contains collections of those data objects (such as a [ChartCollection](/javascript/api/office-scripts/excel/excel.chartcollection)) that are present in the individual sheet.</span></span>
- <span data-ttu-id="c0d0d-115">[ブック](/javascript/api/office-scripts/excel/excel.workbook)。</span><span class="sxs-lookup"><span data-stu-id="c0d0d-115">[Workbooks](/javascript/api/office-scripts/excel/excel.workbook).</span></span> <span data-ttu-id="c0d0d-116">[ブック](/javascript/api/office-scripts/excel/excel.workbook)全体の一部のデータオブジェクト ( [tablecollection](/javascript/api/office-scripts/excel/excel.tablecollection)など) のコレクションが含まれています。</span><span class="sxs-lookup"><span data-stu-id="c0d0d-116">contain collections of some of those data objects (such as a [TableCollection](/javascript/api/office-scripts/excel/excel.tablecollection)) for the entire [Workbook](/javascript/api/office-scripts/excel/excel.workbook).</span></span>

<span data-ttu-id="c0d0d-117">Office Scripts オブジェクトモデルの詳細については、「 [Office スクリプトの基礎知識](/office/dev/scripts/develop/scripting-fundamentals)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c0d0d-117">For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)</span></span>

## <a name="see-also"></a><span data-ttu-id="c0d0d-118">関連項目</span><span class="sxs-lookup"><span data-stu-id="c0d0d-118">See also</span></span>

- [<span data-ttu-id="c0d0d-119">Office スクリプトについて</span><span class="sxs-lookup"><span data-stu-id="c0d0d-119">About Office Scripts</span></span>](/office/dev/scripts/overview/excel)
- [<span data-ttu-id="c0d0d-120">Web 上の Excel で Office スクリプトを記録、編集、および作成する</span><span class="sxs-lookup"><span data-stu-id="c0d0d-120">Record, edit, and create Office Scripts in Excel on the web</span></span>](/office/dev/scripts/tutorials/excel-tutorial)
- [<span data-ttu-id="c0d0d-121">Web 上の Excel での Office スクリプトのスクリプトの基礎</span><span class="sxs-lookup"><span data-stu-id="c0d0d-121">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](/office/dev/scripts/develop/scripting-fundamentals)
