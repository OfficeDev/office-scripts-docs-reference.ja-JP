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
# <a name="office-scripts-api-reference"></a>Office スクリプト API リファレンス

Office スクリプト API を使用すると、web 上の Excel で一般的なタスクを自動化できます。 このリファレンスドキュメントを使用して、スクリプトで使用できるクラス、メソッド、およびその他の種類の詳細について説明します。 Office スクリプトを使用してアクセスできるすべてのオブジェクトについては、ページの左側にある目次を参照してください。

> [!NOTE]
> Office アドインを開発するための JavaScript Api を探している場合は、「 [Office アドインの JAVASCRIPT api リファレンス](/javascript/api/overview?view=excel-js-preview)」を参照してください。

## <a name="common-classes"></a>一般的なクラス

次の一覧は、Office Scripts オブジェクトモデルの基本事項を示しています。 これにより、共通のクラスと、それらが相互にどのように関係しているかがわかります。

- [ブック](/javascript/api/office-scripts/excel/excelscript.workbook) には、1 つ以上の [ワークシート](/javascript/api/office-scripts/excel/excelscript.worksheet) が含まれます。
- [ワークシート](/javascript/api/office-scripts/excel/excelscript.worksheet) では、[Range](/javascript/api/office-scripts/excel/excelscript.range) オブジェクトを介してセルにアクセスできます。
- [Range](/javascript/api/office-scripts/excel/excelscript.range) は、連続したセルのグループを表します。
- [Range](/javascript/api/office-scripts/excel/excelscript.range) は、[表](/javascript/api/office-scripts/excel/excelscript.table)、[グラフ](/javascript/api/office-scripts/excel/excelscript.chart)、[図形](/javascript/api/office-scripts/excel/excelscript.shape)、およびその他のデータ可視化や組織オブジェクトを作成して配置するために使用されます。
- [ワークシート](/javascript/api/office-scripts/excel/excelscript.worksheet)には、個々のシートにあるオブジェクトが格納されている配列が含まれています。
- ブックには、ブック全体に対するいくつかのデータオブジェクトの配列[が含まれ](/javascript/api/office-scripts/excel/excelscript.workbook)ています。

Office Scripts オブジェクトモデルの詳細については、「 [Office スクリプトの基礎知識](/office/dev/scripts/develop/scripting-fundamentals)」を参照してください。

## <a name="see-also"></a>関連項目

- [Office スクリプトについて](/office/dev/scripts/overview/excel)
- [Excel on the web で Office スクリプトを記録、編集、作成する](/office/dev/scripts/tutorials/excel-tutorial)
- [Excel on the web での Office スクリプトのスクリプトの基本事項](/office/dev/scripts/develop/scripting-fundamentals)
