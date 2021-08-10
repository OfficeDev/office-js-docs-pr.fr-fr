---
title: Étendre des fonctions personnalisées avec des fonctions XLL définies par l’utilisateur
description: Activer la compatibilité avec Excel fonctions XLL définies par l’utilisateur qui ont des fonctionnalités équivalentes à vos fonctions personnalisées
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: b7a2330f7a875c894f371138034314ae99bb0e9393a45c6e8572a97a084fe94e
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089312"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a>Étendre des fonctions personnalisées avec des fonctions XLL définies par l’utilisateur

Si vous avez des XL Excel existantes, vous pouvez créer des fonctions personnalisées équivalentes dans un module Excel pour étendre les fonctionnalités de votre solution à d’autres plateformes telles que en ligne ou sur un Mac. Toutefois, les Excel ne disposent pas de toutes les fonctionnalités disponibles dans les XL. Selon les fonctionnalités que votre solution utilise, la XLL peut offrir une meilleure expérience que les fonctions personnalisées du Excel dans Excel sur Windows.

> [!NOTE]
> La compatibilité des UDF et des Microsoft 365 COM est prise en charge par les plateformes suivantes.
>
> - Excel sur le web
> - Excel sur Windows (version 1904 ou ultérieure)
> - Excel sur Mac (version 13.329 ou ultérieure)
>
> Pour utiliser la compatibilité des UDF et du Excel sur le Web COM, connectez-vous à l’aide de votre abonnement Microsoft 365 ou [d’un compte Microsoft.](https://account.microsoft.com/account) Si vous n’avez pas encore d’abonnement Microsoft 365, vous pouvez obtenir un abonnement Microsoft 365 de 90 jours renouvelable gratuit en rejoignant le programme Microsoft 365 [développeur.](https://developer.microsoft.com/office/dev-program)

## <a name="specify-equivalent-xll-in-the-manifest"></a>Spécifier un XLL équivalent dans le manifeste

Pour activer la compatibilité avec un XLL existant, identifiez le XLL équivalent dans le manifeste de votre Excel de données. Excel utilisera ensuite les fonctions XLL au lieu de vos fonctions personnalisées de Excel lors de l’exécution sur Windows.

Pour définir le XLL équivalent pour vos fonctions personnalisées, spécifiez `FileName` le XLL. Lorsque l’utilisateur ouvre un workbook avec des fonctions du XLL, Excel convertit les fonctions en fonctions compatibles. Le workbook utilise ensuite le XLL lorsqu’il est ouvert dans Excel sur Windows et utilise des fonctions personnalisées à partir de votre Excel lorsqu’il est ouvert en ligne ou sur un Mac.

L’exemple suivant montre comment spécifier un add-in COM et un XLL comme équivalent. Souvent, vous spécifiez les deux. Pour plus d’complétance, cet exemple montre les deux en contexte. Ils sont identifiés par `ProgId` leur `FileName` et, respectivement. `EquivalentAddins`L’élément doit être placé immédiatement avant la balise de `VersionOverrides` fermeture. Pour plus d’informations sur la compatibilité des applications COM, voir Rendre votre Office compatible avec un compl?ment [COM existant.](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)

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
> Si un add-in déclare que ses fonctions personnalisées sont compatibles avec XLL, la modification ultérieure du manifeste peut rompre le classez d’un utilisateur, car il modifiera le format de fichier.

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>Comportement des fonctions personnalisées pour les fonctions compatibles XLL

Les fonctions XLL d’un add-in sont converties en fonctions personnalisées compatibles avec XLL lorsqu’une feuille de calcul est ouverte et qu’un module équivalent est disponible. Lors de l’enregistrer suivant, les fonctions XLL sont écrites dans le fichier dans un mode compatible afin qu’elles fonctionnent avec les fonctions personnalisées de la XLL et du Excel (sur d’autres plateformes).

Le tableau suivant compare les fonctionnalités entre les fonctions XLL définies par l’utilisateur, les fonctions personnalisées compatibles XLL et Excel fonctions personnalisées de modules.

|         |Fonction XLL définie par l’utilisateur |Fonctions personnalisées compatibles XLL |Excel fonction personnalisée de add-in |
|---------|---------|---------|---------|
| **Plateformes prises en charge** | Windows | Windows, macOS, navigateur web | Windows, macOS, navigateur web |
| **Formats de fichiers pris en charge** | XLSX, XLSB, XLSM, XLS | XLSX, XLSB, XLSM | XLSX, XLSB, XLSM |
| **Autocomplete de formule** | Non | Oui | Oui |
| **Diffusion en continu** | Possible via le rappel xlfRTD et XLL. | Oui | Oui |
| **Localisation des fonctions** | Non | Non. Le nom et l’ID doivent correspondre aux fonctions XLL existantes. | Oui |
| **Fonctions volatiles** | Oui | Oui | Oui |
| **Prise en charge du recalcul multi-thread** | Oui | Oui | Oui |
| **Comportement du calcul** | Aucune interface utilisateur. Excel ne répond pas pendant le calcul. | Les utilisateurs voient #BUSY ! jusqu’à ce qu’un résultat soit renvoyé. | Les utilisateurs voient #BUSY ! jusqu’à ce qu’un résultat soit renvoyé. |
| **Ensembles de conditions requises** | N/A | CustomFunctions 1.1 et les ultérieures | CustomFunctions 1.1 et les ultérieures |

## <a name="see-also"></a>Voir aussi

- [Rendre votre complément Office compatible avec un complément COM existant](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
