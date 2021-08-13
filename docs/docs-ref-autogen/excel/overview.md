---
title: Office スクリプト API リファレンス
description: JavaScript API Officeの概要。
ms.date: 06/29/2020
ms.openlocfilehash: 3ce3344fb49b2811719feb13f8fb4118f1a20060db9d85a06d1be939f22bf3c5
ms.sourcegitcommit: 6a0182075a558c4fe664fedfee08fea76513b192
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2021
ms.locfileid: "58183427"
---
# <a name="office-scripts-api-reference"></a>Office スクリプト API リファレンス

このOfficeスクリプト API を使用すると、一般的なタスクを自動化Excel on the web。 スクリプトで使用できるクラス、メソッド、その他の種類の詳細については、このリファレンス ドキュメントを参照してください。 スクリプトからアクセスOfficeオブジェクトはすべて、ページの左側の目次に表示されます。

> [!NOTE]
> アドインを開発するための JavaScript API を探している場合Officeアドインの[JavaScript API](/javascript/api/overview?view=excel-js-preview)リファレンスOffice参照してください。

## <a name="common-classes"></a>共通クラス

次の一覧では、Scripts オブジェクト モデルのOfficeを示します。 これは、共通のクラスと、そのクラスが他のクラスとどのように関連付け合うのかを示しています。

- [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook) には、1 つ以上の [ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet) が含まれます。
- [ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet) では、[Range](/javascript/api/office-scripts/excelscript/excelscript.range) オブジェクトを介してセルにアクセスできます。
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range) は、連続したセルのグループを表します。
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range) は、[表](/javascript/api/office-scripts/excelscript/excelscript.table)、[グラフ](/javascript/api/office-scripts/excelscript/excelscript.chart)、[図形](/javascript/api/office-scripts/excelscript/excelscript.shape)、およびその他のデータ可視化や組織オブジェクトを作成して配置するために使用されます。
- Worksheet [には](/javascript/api/office-scripts/excelscript/excelscript.worksheet) 、個々のシートに存在するオブジェクトで埋め込まれている配列が含まれます。
- ブック [には](/javascript/api/office-scripts/excelscript/excelscript.workbook) 、ブック全体のデータ オブジェクトの一部の配列が含まれます。

スクリプト オブジェクト モデルの詳細については、「Office スクリプトのスクリプトの基本[Office」を参照Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>関連項目

- [スクリプトOfficeについて](/office/dev/scripts/overview/excel)
- [Excel on the web で Office スクリプトを記録、編集、作成する](/office/dev/scripts/tutorials/excel-tutorial)
- [Excel on the web での Office スクリプトのスクリプトの基本事項](/office/dev/scripts/develop/scripting-fundamentals)
