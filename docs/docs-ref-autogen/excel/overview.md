---
title: Office スクリプト API リファレンス
description: Office Scripts JavaScript Api の概要について説明します。
ms.date: 06/29/2020
ms.openlocfilehash: 7c4fe97ca35cfb442ebbf9db2e0b03b389185ae8
ms.sourcegitcommit: 9c4c4c213a203e58c55eb3d84d7d92fa527f3eb8
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/01/2020
ms.locfileid: "45004732"
---
# <a name="office-scripts-api-reference"></a><span data-ttu-id="1f157-103">Office スクリプト API リファレンス</span><span class="sxs-lookup"><span data-stu-id="1f157-103">Office Scripts API reference</span></span>

<span data-ttu-id="1f157-104">Office スクリプト API を使用すると、web 上の Excel で一般的なタスクを自動化できます。</span><span class="sxs-lookup"><span data-stu-id="1f157-104">The Office Scripts API lets you automate common tasks in Excel on the web.</span></span> <span data-ttu-id="1f157-105">このリファレンスドキュメントを使用して、スクリプトで使用できるクラス、メソッド、およびその他の種類の詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="1f157-105">Use this reference documentation to learn more about the classes, methods, and other types available for your scripts.</span></span> <span data-ttu-id="1f157-106">Office スクリプトを使用してアクセスできるすべてのオブジェクトについては、ページの左側にある目次を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1f157-106">All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.</span></span>

> [!NOTE]
> <span data-ttu-id="1f157-107">Office アドインを開発するための JavaScript Api を探している場合は、「 [Office アドインの JAVASCRIPT api リファレンス](/javascript/api/overview?view=excel-js-preview)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1f157-107">If you're looking for the JavaScript APIs for developing Office Add-ins, visit the [Office Add-ins JavaScript API reference](/javascript/api/overview?view=excel-js-preview).</span></span>

## <a name="common-classes"></a><span data-ttu-id="1f157-108">一般的なクラス</span><span class="sxs-lookup"><span data-stu-id="1f157-108">Common classes</span></span>

<span data-ttu-id="1f157-109">次の一覧は、Office Scripts オブジェクトモデルの基本事項を示しています。</span><span class="sxs-lookup"><span data-stu-id="1f157-109">The following list breaks down the basics of the Office Scripts object model.</span></span> <span data-ttu-id="1f157-110">これにより、共通のクラスと、それらが相互にどのように関係しているかがわかります。</span><span class="sxs-lookup"><span data-stu-id="1f157-110">This shows the common classes and how they relate to one another.</span></span>

- <span data-ttu-id="1f157-111">[ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook) には、1 つ以上の [ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet) が含まれます。</span><span class="sxs-lookup"><span data-stu-id="1f157-111">A [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excelscript/excelscript.worksheet).</span></span>
- <span data-ttu-id="1f157-112">[ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet) では、[Range](/javascript/api/office-scripts/excelscript/excelscript.range) オブジェクトを介してセルにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="1f157-112">A [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excelscript/excelscript.range) objects.</span></span>
- <span data-ttu-id="1f157-113">[Range](/javascript/api/office-scripts/excelscript/excelscript.range) は、連続したセルのグループを表します。</span><span class="sxs-lookup"><span data-stu-id="1f157-113">A [Range](/javascript/api/office-scripts/excelscript/excelscript.range) represents a group of contiguous cells.</span></span>
- <span data-ttu-id="1f157-114">[Range](/javascript/api/office-scripts/excelscript/excelscript.range) は、[表](/javascript/api/office-scripts/excelscript/excelscript.table)、[グラフ](/javascript/api/office-scripts/excelscript/excelscript.chart)、[図形](/javascript/api/office-scripts/excelscript/excelscript.shape)、およびその他のデータ可視化や組織オブジェクトを作成して配置するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="1f157-114">[Ranges](/javascript/api/office-scripts/excelscript/excelscript.range) are used to create and place [Tables](/javascript/api/office-scripts/excelscript/excelscript.table), [Charts](/javascript/api/office-scripts/excelscript/excelscript.chart), [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape), and other data visualization or organization objects.</span></span>
- <span data-ttu-id="1f157-115">[ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet)には、個々のシートにあるオブジェクトが格納されている配列が含まれています。</span><span class="sxs-lookup"><span data-stu-id="1f157-115">A [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) contains arrays filled with those objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="1f157-116">ブックには、ブック全体に対するいくつかのデータオブジェクトの配列[が含まれ](/javascript/api/office-scripts/excelscript/excelscript.workbook)ています。</span><span class="sxs-lookup"><span data-stu-id="1f157-116">A [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) contains arrays of some of those data objects for the entire Workbook.</span></span>

<span data-ttu-id="1f157-117">Office Scripts オブジェクトモデルの詳細については、「 [Office スクリプトの基礎知識](/office/dev/scripts/develop/scripting-fundamentals)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1f157-117">For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)</span></span>

## <a name="see-also"></a><span data-ttu-id="1f157-118">関連項目</span><span class="sxs-lookup"><span data-stu-id="1f157-118">See also</span></span>

- [<span data-ttu-id="1f157-119">Office スクリプトについて</span><span class="sxs-lookup"><span data-stu-id="1f157-119">About Office Scripts</span></span>](/office/dev/scripts/overview/excel)
- [<span data-ttu-id="1f157-120">Excel on the web で Office スクリプトを記録、編集、作成する</span><span class="sxs-lookup"><span data-stu-id="1f157-120">Record, edit, and create Office Scripts in Excel on the web</span></span>](/office/dev/scripts/tutorials/excel-tutorial)
- [<span data-ttu-id="1f157-121">Excel on the web での Office スクリプトのスクリプトの基本事項</span><span class="sxs-lookup"><span data-stu-id="1f157-121">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](/office/dev/scripts/develop/scripting-fundamentals)
