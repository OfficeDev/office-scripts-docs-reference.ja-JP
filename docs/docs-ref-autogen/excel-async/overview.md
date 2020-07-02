---
title: Office スクリプト API リファレンス
description: Office スクリプトの非同期 JavaScript Api の概要。
ms.date: 06/29/2020
ms.openlocfilehash: 3d8d37b30d9535e8b6a56a08c44f9034cb599f31
ms.sourcegitcommit: 9c4c4c213a203e58c55eb3d84d7d92fa527f3eb8
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/01/2020
ms.locfileid: "45003955"
---
# <a name="office-scripts-async-api-reference"></a>Office スクリプトの非同期 API リファレンス

Office Scripts Async API は、Office スクリプトのプレビュー段階で作成された古いスクリプトをサポートしています。 これらの以前のスクリプトで使用されていたクラス、メソッド、およびその他の種類の詳細については、このリファレンスドキュメントを使用してください。 Office スクリプトを使用してアクセスできるすべてのオブジェクトについては、ページの左側にある目次を参照してください。

> [!IMPORTANT]
> 標準の Office スクリプト Api を使用して、新しいスクリプトを作成することを強くお勧めします。 新しいスクリプトを作成または編集している場合は、[非非同期バージョン](?view=office-scripts)の api に切り替えてください。

## <a name="common-classes"></a>一般的なクラス

次の一覧は、Office Scripts オブジェクトモデルの基本事項を示しています。 これにより、共通のクラスと、それらが相互にどのように関係しているかがわかります。

- [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook)には、1つまたは複数の[ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet)が含まれて[います。](/javascript/api/office-scripts/excelscript/excelscript.worksheetcollection)
- [ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet) では、[Range](/javascript/api/office-scripts/excelscript/excelscript.range) オブジェクトを介してセルにアクセスできます。
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range) は、連続したセルのグループを表します。
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range) は、[表](/javascript/api/office-scripts/excelscript/excelscript.table)、[グラフ](/javascript/api/office-scripts/excelscript/excelscript.chart)、[図形](/javascript/api/office-scripts/excelscript/excelscript.shape)、およびその他のデータ可視化や組織オブジェクトを作成して配置するために使用されます。
- [ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet)には、個々のシートにあるデータオブジェクト ( [chartcollection](/javascript/api/office-scripts/excelscript/excelscript.chartcollection)など) のコレクションが含まれています。
- [ブック全体](/javascript/api/office-scripts/excelscript/excelscript.workbook)の一部のデータオブジェクト ( [tablecollection](/javascript/api/office-scripts/excelscript/excelscript.tablecollection)など) のコレクションが含まれ[ています](/javascript/api/office-scripts/excelscript/excelscript.workbook)。

Office Scripts オブジェクトモデルの詳細については、「 [Office スクリプトの基礎知識](/office/dev/scripts/develop/scripting-fundamentals)」を参照してください。

## <a name="see-also"></a>関連項目

- [従来のスクリプトをサポートするための Office スクリプト非同期 Api の使用](/office/dev/scripts/develop/excel-async-model)
- [Office スクリプトについて](/office/dev/scripts/overview/excel)
- [Excel on the web で Office スクリプトを記録、編集、作成する](/office/dev/scripts/tutorials/excel-tutorial)
- [Excel on the web での Office スクリプトのスクリプトの基本事項](/office/dev/scripts/develop/scripting-fundamentals)
