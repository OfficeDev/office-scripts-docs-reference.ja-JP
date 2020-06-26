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
# <a name="office-scripts-async-api-reference"></a>Office スクリプトの非同期 API リファレンス

Office Scripts Async API は、Office スクリプトのプレビュー段階で作成された古いスクリプトをサポートしています。 これらの以前のスクリプトで使用されていたクラス、メソッド、およびその他の種類の詳細については、このリファレンスドキュメントを使用してください。 Office スクリプトを使用してアクセスできるすべてのオブジェクトについては、ページの左側にある目次を参照してください。

> [!IMPORTANT]
> 標準の Office スクリプト Api を使用して、新しいスクリプトを作成することを強くお勧めします。 新しいスクリプトを作成または編集している場合は、[非非同期バージョン](?view=office-scripts)の api に切り替えてください。

## <a name="common-classes"></a>一般的なクラス

次の一覧は、Office Scripts オブジェクトモデルの基本事項を示しています。 これにより、共通のクラスと、それらが相互にどのように関係しているかがわかります。

- [ブック](/javascript/api/office-scripts/excel/excelscript.workbook)には、1つまたは複数の[ワークシート](/javascript/api/office-scripts/excel/excelscript.worksheet)が含まれて[います。](/javascript/api/office-scripts/excel/excelscript.worksheetcollection)
- [ワークシート](/javascript/api/office-scripts/excel/excelscript.worksheet) では、[Range](/javascript/api/office-scripts/excel/excelscript.range) オブジェクトを介してセルにアクセスできます。
- [Range](/javascript/api/office-scripts/excel/excelscript.range) は、連続したセルのグループを表します。
- [Range](/javascript/api/office-scripts/excel/excelscript.range) は、[表](/javascript/api/office-scripts/excel/excelscript.table)、[グラフ](/javascript/api/office-scripts/excel/excelscript.chart)、[図形](/javascript/api/office-scripts/excel/excelscript.shape)、およびその他のデータ可視化や組織オブジェクトを作成して配置するために使用されます。
- [ワークシート](/javascript/api/office-scripts/excel/excelscript.worksheet)には、個々のシートにあるデータオブジェクト ( [chartcollection](/javascript/api/office-scripts/excel/excelscript.chartcollection)など) のコレクションが含まれています。
- [ブック全体](/javascript/api/office-scripts/excel/excelscript.workbook)の一部のデータオブジェクト ( [tablecollection](/javascript/api/office-scripts/excel/excelscript.tablecollection)など) のコレクションが含まれ[ています](/javascript/api/office-scripts/excel/excelscript.workbook)。

Office Scripts オブジェクトモデルの詳細については、「 [Office スクリプトの基礎知識](/office/dev/scripts/develop/scripting-fundamentals)」を参照してください。

## <a name="see-also"></a>関連項目

- [従来のスクリプトをサポートするための Office スクリプト非同期 Api の使用](/office/dev/scripts/develop/excel-async-model)
- [Office スクリプトについて](/office/dev/scripts/overview/excel)
- [Excel on the web で Office スクリプトを記録、編集、作成する](/office/dev/scripts/tutorials/excel-tutorial)
- [Excel on the web での Office スクリプトのスクリプトの基本事項](/office/dev/scripts/develop/scripting-fundamentals)
