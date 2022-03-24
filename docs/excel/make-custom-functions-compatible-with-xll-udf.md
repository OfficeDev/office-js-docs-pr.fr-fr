---
title: Étendre des fonctions personnalisées avec des fonctions XLL définies par l’utilisateur
description: Activez la compatibilité avec Excel fonctions XLL définies par l’utilisateur qui ont des fonctionnalités équivalentes à vos fonctions personnalisées.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: dac6cdceb65f27c7246afe17721ba4d11bbf18ab
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745649"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a>Étendre des fonctions personnalisées avec des fonctions XLL définies par l’utilisateur

> [!NOTE]
> Un Excel XLL est un fichier de Excel avec l’extension **de fichier .xll**. Un fichier XLL est un type de fichier de bibliothèque de liens dynamiques (DLL) qui ne peut être ouvert qu’Excel. Les fichiers de add-in XLL doivent être écrits en C ou C++. Pour [en savoir plus, Excel développement de XLS](/office/client-developer/excel/developing-excel-xlls).

Si vous disposez de Excel XLL, vous pouvez créer des macros supplémentaires de fonction personnalisée équivalentes à l’aide de l’API JavaScript Excel pour étendre vos fonctionnalités de solution à d’autres plateformes, telles que Excel sur le Web ou sur un Mac. Toutefois, Excel’API JavaScript ne disposent pas de toutes les fonctionnalités disponibles dans les add-ins XLL. En fonction des fonctionnalités que votre solution utilise, le add-in XLL peut offrir une meilleure expérience que le Excel de l’API JavaScript dans Excel sur Windows.

[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

## <a name="specify-equivalent-xll-in-the-manifest"></a>Spécifier un XLL équivalent dans le manifeste

Pour activer la compatibilité avec un compl?ment XLL existant, identifiez le compl?ment XLL équivalent dans le manifeste de votre compl?ment d’API JavaScript Excel. Excel utilisera ensuite les fonctions du add-in XLL au lieu de vos fonctions personnalisées d’API JavaScript Excel lors de l’exécution sur Windows.

Pour définir le modèle XLL équivalent pour vos fonctions personnalisées, spécifiez le `FileName` fichier XLL. Lorsque l’utilisateur ouvre un classez avec des fonctions à partir du fichier XLL, Excel convertit les fonctions en fonctions compatibles. Le classez utilise ensuite le fichier XLL lorsqu’il est ouvert dans Excel sur Windows et utilise des fonctions personnalisées à partir de votre add-in d’API JavaScript Excel lorsqu’il est ouvert sur le web ou sur un Mac.

L’exemple suivant montre comment spécifier un compl?ment COM et un compl?ment XLL en tant qu’équivalents dans un fichier manifeste de l’API JavaScript Excel. Souvent, vous spécifiez les deux. Pour plus d’complétance, cet exemple montre les deux en contexte. Ils sont identifiés par leur et `ProgId` `FileName` , respectivement. L’élément `EquivalentAddins` doit être placé immédiatement avant la balise de `VersionOverrides` fermeture. Pour plus d’informations sur la compatibilité des applications COM, voir Rendre votre Office compatible avec un compl?ment [COM existant](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>

    <EquivalentAddin>
      <FileName>contosofunctions.xll</FileName>
      <Type>XLL</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

> [!NOTE]
> Si un Excel d’API JavaScript déclare que ses fonctions personnalisées sont compatibles avec un add-in XLL, la modification ultérieure du manifeste peut rompre le classez d’un utilisateur, car il modifiera le format de fichier.

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>Comportement des fonctions personnalisées pour les fonctions compatibles XLL

Les fonctions XLL d’un add-in sont converties en fonctions personnalisées compatibles avec XLL lorsqu’une feuille de calcul est ouverte et qu’un module équivalent est disponible. Lors de l’enregistrer suivant, les fonctions XLL sont écrites dans le fichier dans un mode compatible afin qu’elles fonctionnent avec les fonctions personnalisées de l’API JavaScript et du add-in XLL Excel (sur d’autres plateformes).

Le tableau suivant compare les fonctionnalités entre les fonctions XLL définies par l’utilisateur, les fonctions personnalisées compatibles XLL et Excel fonctions personnalisées de l’API JavaScript.

|         |Fonction XLL définie par l’utilisateur |Fonctions personnalisées compatibles XLL |Excel fonction personnalisée de l’API JavaScript |
|---------|---------|---------|---------|
| **Plateformes prises en charge** | Windows | Windows, macOS, navigateur web | Windows, macOS, navigateur web |
| **Formats de fichiers pris en charge** | XLSX, XLSB, XLSM, XLS | XLSX, XLSB, XLSM | XLSX, XLSB, XLSM |
| **Autocomplete de formule** | Non | Oui | Oui |
| **Diffusion en continu** | Possible via le rappel xlfRTD et XLL. | Oui | Oui |
| **Localisation des fonctions** | Non | Non. Le nom et l’ID doivent correspondre aux fonctions XLL existantes. | Oui |
| **Fonctions volatiles** | Oui | Oui | Oui |
| **Prise en charge du recalcul multi-thread** | Oui | Oui | Oui |
| **Comportement du calcul** | Aucune interface utilisateur. Excel ne répond pas pendant le calcul. | Les utilisateurs voient #BUSY ! jusqu’à ce qu’un résultat soit renvoyé. | Les utilisateurs voient #BUSY ! jusqu’à ce qu’un résultat soit renvoyé. |
| **Ensembles de conditions requises** | S/O | CustomFunctions 1.1 et les ultérieures | CustomFunctions 1.1 et les ultérieures |

## <a name="see-also"></a>Voir aussi

- [Rendre votre complément Office compatible avec un complément COM existant](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
