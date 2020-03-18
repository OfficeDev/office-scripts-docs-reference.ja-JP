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
# <a name="office-scripts-api-reference"></a>Office スクリプト API リファレンス

Office スクリプト API を使用すると、web 上の Excel で一般的なタスクを自動化できます。 このリファレンスドキュメントを使用して、スクリプトで使用できるクラス、メソッド、およびその他の種類の詳細について説明します。 Office スクリプトを使用してアクセスできるすべてのオブジェクトについては、ページの左側にある目次を参照してください。

## <a name="common-classes"></a>一般的なクラス

次の一覧は、Office Scripts オブジェクトモデルの基本事項を示しています。 これにより、共通のクラスと、それらが相互にどのように関係しているかがわかります。

- [ブック](/javascript/api/office-scripts/excel/excel.workbook)には、1つまたは複数の[ワークシート](/javascript/api/office-scripts/excel/excel.worksheet)が含まれて[います。](/javascript/api/office-scripts/excel/excel.worksheetcollection)
- [ワークシート](/javascript/api/office-scripts/excel/excel.worksheet)は、 [Range](/javascript/api/office-scripts/excel/excel.range)オブジェクトを通じてセルへのアクセスを提供します。
- [範囲](/javascript/api/office-scripts/excel/excel.range)は、連続したセルのグループを表します。
- [範囲](/javascript/api/office-scripts/excel/excel.range)は、[表](/javascript/api/office-scripts/excel/excel.table)、[グラフ](/javascript/api/office-scripts/excel/excel.chart)、[図形](/javascript/api/office-scripts/excel/excel.shape)、およびその他のデータビジュアライゼーションまたは組織オブジェクトを作成して配置するために使用されます。
- [ワークシート](/javascript/api/office-scripts/excel/excel.worksheet)には、個々のシートにあるデータオブジェクト ( [chartcollection](/javascript/api/office-scripts/excel/excel.chartcollection)など) のコレクションが含まれています。
- [ブック](/javascript/api/office-scripts/excel/excel.workbook)。 [ブック](/javascript/api/office-scripts/excel/excel.workbook)全体の一部のデータオブジェクト ( [tablecollection](/javascript/api/office-scripts/excel/excel.tablecollection)など) のコレクションが含まれています。

Office Scripts オブジェクトモデルの詳細については、「 [Office スクリプトの基礎知識](/office/dev/scripts/develop/scripting-fundamentals)」を参照してください。

## <a name="see-also"></a>関連項目

- [Office スクリプトについて](/office/dev/scripts/overview/excel)
- [Web 上の Excel で Office スクリプトを記録、編集、および作成する](/office/dev/scripts/tutorials/excel-tutorial)
- [Web 上の Excel での Office スクリプトのスクリプトの基礎](/office/dev/scripts/develop/scripting-fundamentals)
