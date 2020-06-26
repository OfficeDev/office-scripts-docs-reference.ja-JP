---
title: Office スクリプト API リファレンス
description: Office Scripts JavaScript Api の概要について説明します。
ms.date: 06/17/2020
ms.openlocfilehash: 5634d0e5f68464655054ad1c09eb7931e0da62d4
ms.sourcegitcommit: 163b26a43411ad7f13a01237efe9b8d6de656b47
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44884838"
---
# <a name="office-scripts-api-reference"></a><span data-ttu-id="8b6d3-103">Office スクリプト API リファレンス</span><span class="sxs-lookup"><span data-stu-id="8b6d3-103">Office Scripts API reference</span></span>

<span data-ttu-id="8b6d3-104">Office スクリプト API を使用すると、web 上の Excel で一般的なタスクを自動化できます。</span><span class="sxs-lookup"><span data-stu-id="8b6d3-104">The Office Scripts API lets you automate common tasks in Excel on the web.</span></span> <span data-ttu-id="8b6d3-105">このリファレンスドキュメントを使用して、スクリプトで使用できるクラス、メソッド、およびその他の種類の詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="8b6d3-105">Use this reference documentation to learn more about the classes, methods, and other types available for your scripts.</span></span> <span data-ttu-id="8b6d3-106">Office スクリプトを使用してアクセスできるすべてのオブジェクトについては、ページの左側にある目次を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8b6d3-106">All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.</span></span>

> [!NOTE]
> <span data-ttu-id="8b6d3-107">Office アドインを開発するための JavaScript Api を探している場合は、「 [Office アドインの JAVASCRIPT api リファレンス](/javascript/api/overview?view=excel-js-preview)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8b6d3-107">If you're looking for the JavaScript APIs for developing Office Add-ins, visit the [Office Add-ins JavaScript API reference](/javascript/api/overview?view=excel-js-preview).</span></span>

## <a name="common-classes"></a><span data-ttu-id="8b6d3-108">一般的なクラス</span><span class="sxs-lookup"><span data-stu-id="8b6d3-108">Common classes</span></span>

<span data-ttu-id="8b6d3-109">次の一覧は、Office Scripts オブジェクトモデルの基本事項を示しています。</span><span class="sxs-lookup"><span data-stu-id="8b6d3-109">The following list breaks down the basics of the Office Scripts object model.</span></span> <span data-ttu-id="8b6d3-110">これにより、共通のクラスと、それらが相互にどのように関係しているかがわかります。</span><span class="sxs-lookup"><span data-stu-id="8b6d3-110">This shows the common classes and how they relate to one another.</span></span>

- <span data-ttu-id="8b6d3-111">[ブック](/javascript/api/office-scripts/excel/excelscript.workbook) には、1 つ以上の [ワークシート](/javascript/api/office-scripts/excel/excelscript.worksheet) が含まれます。</span><span class="sxs-lookup"><span data-stu-id="8b6d3-111">A [Workbook](/javascript/api/office-scripts/excel/excelscript.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excel/excelscript.worksheet).</span></span>
- <span data-ttu-id="8b6d3-112">[ワークシート](/javascript/api/office-scripts/excel/excelscript.worksheet) では、[Range](/javascript/api/office-scripts/excel/excelscript.range) オブジェクトを介してセルにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="8b6d3-112">A [Worksheet](/javascript/api/office-scripts/excel/excelscript.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excel/excelscript.range) objects.</span></span>
- <span data-ttu-id="8b6d3-113">[Range](/javascript/api/office-scripts/excel/excelscript.range) は、連続したセルのグループを表します。</span><span class="sxs-lookup"><span data-stu-id="8b6d3-113">A [Range](/javascript/api/office-scripts/excel/excelscript.range) represents a group of contiguous cells.</span></span>
- <span data-ttu-id="8b6d3-114">[Range](/javascript/api/office-scripts/excel/excelscript.range) は、[表](/javascript/api/office-scripts/excel/excelscript.table)、[グラフ](/javascript/api/office-scripts/excel/excelscript.chart)、[図形](/javascript/api/office-scripts/excel/excelscript.shape)、およびその他のデータ可視化や組織オブジェクトを作成して配置するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="8b6d3-114">[Ranges](/javascript/api/office-scripts/excel/excelscript.range) are used to create and place [Tables](/javascript/api/office-scripts/excel/excelscript.table), [Charts](/javascript/api/office-scripts/excel/excelscript.chart), [Shapes](/javascript/api/office-scripts/excel/excelscript.shape), and other data visualization or organization objects.</span></span>
- <span data-ttu-id="8b6d3-115">[ワークシート](/javascript/api/office-scripts/excel/excelscript.worksheet)には、個々のシートにあるオブジェクトが格納されている配列が含まれています。</span><span class="sxs-lookup"><span data-stu-id="8b6d3-115">A [Worksheet](/javascript/api/office-scripts/excel/excelscript.worksheet) contains arrays filled with those objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="8b6d3-116">ブックには、ブック全体に対するいくつかのデータオブジェクトの配列[が含まれ](/javascript/api/office-scripts/excel/excelscript.workbook)ています。</span><span class="sxs-lookup"><span data-stu-id="8b6d3-116">A [Workbook](/javascript/api/office-scripts/excel/excelscript.workbook) contains arrays of some of those data objects for the entire Workbook.</span></span>

<span data-ttu-id="8b6d3-117">Office Scripts オブジェクトモデルの詳細については、「 [Office スクリプトの基礎知識](/office/dev/scripts/develop/scripting-fundamentals)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8b6d3-117">For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)</span></span>

## <a name="see-also"></a><span data-ttu-id="8b6d3-118">関連項目</span><span class="sxs-lookup"><span data-stu-id="8b6d3-118">See also</span></span>

- [<span data-ttu-id="8b6d3-119">Office スクリプトについて</span><span class="sxs-lookup"><span data-stu-id="8b6d3-119">About Office Scripts</span></span>](/office/dev/scripts/overview/excel)
- [<span data-ttu-id="8b6d3-120">Excel on the web で Office スクリプトを記録、編集、作成する</span><span class="sxs-lookup"><span data-stu-id="8b6d3-120">Record, edit, and create Office Scripts in Excel on the web</span></span>](/office/dev/scripts/tutorials/excel-tutorial)
- [<span data-ttu-id="8b6d3-121">Excel on the web での Office スクリプトのスクリプトの基本事項</span><span class="sxs-lookup"><span data-stu-id="8b6d3-121">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](/office/dev/scripts/develop/scripting-fundamentals)
