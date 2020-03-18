---
title: Étendre des fonctions personnalisées avec des fonctions XLL définies par l’utilisateur
description: Activer la compatibilité avec les fonctions Excel XLL définies par l’utilisateur qui offrent une fonctionnalité équivalente à vos fonctions personnalisées
ms.date: 07/31/2019
localization_priority: Normal
ms.openlocfilehash: 8c83ea8a5b97cd4f23a9e25fddabe3d178b1e482
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717039"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a>Étendre des fonctions personnalisées avec des fonctions XLL définies par l’utilisateur

Si vous avez des XLL Excel existantes, vous pouvez créer des fonctions personnalisées équivalentes dans un complément Excel pour étendre les fonctionnalités de votre solution à d’autres plateformes, comme Online ou macOS. Toutefois, les compléments Excel ne disposent pas de toutes les fonctionnalités disponibles dans les XLL. En fonction de la fonctionnalité utilisée par votre solution, le XLL peut offrir une meilleure expérience que les fonctions personnalisées de complément Excel dans Excel sur Windows.

> [!NOTE]
> Le complément COM et la compatibilité UDF XLL sont pris en charge par les plateformes suivantes, lorsqu’ils sont connectés à un abonnement Office 365 :
> - Excel sur le web
> - Excel sur Windows (version 1904 ou ultérieure)
> - Excel sur Mac (version 13,329 ou ultérieure)
> 
> Pour utiliser le complément COM et la compatibilité des FDU XLL dans Excel sur le Web, connectez-vous à l’aide de votre abonnement Office 365 ou d’un [compte Microsoft](https://account.microsoft.com/account). Si vous ne disposez pas déjà d’un abonnement Office 365, vous pouvez obtenir gratuitement un abonnement Office 365 renouvelable 90 jours en joignant le [programme de développement office 365](https://developer.microsoft.com/office/dev-program).

## <a name="specify-equivalent-xll-in-the-manifest"></a>Spécifier le XLL équivalent dans le manifeste

Pour activer la compatibilité avec un XLL existant, identifiez le XLL équivalent dans le manifeste de votre complément Excel. Ensuite, Excel utilise les fonctions de la XLL au lieu de vos fonctions personnalisées de complément Excel lors de l’exécution de Windows.

Pour définir le XLL équivalent pour vos fonctions personnalisées, spécifiez l' `FileName` élément XLL. Lorsque l’utilisateur ouvre un classeur avec des fonctions à partir de la XLL, Excel convertit les fonctions en fonctions compatibles. Le classeur utilise ensuite le XLL lorsqu’il est ouvert dans Excel sur Windows et utilise des fonctions personnalisées à partir de votre complément Excel lorsqu’il est ouvert en ligne ou sur macOS.

L’exemple suivant montre comment spécifier un complément COM et un XLL comme équivalent. Souvent, vous spécifierez à la fois de manière à ce que cet exemple montre les deux dans le contexte. Ils sont identifiés par leur `ProgId` et `FileName` respectivement. L' `EquivalentAddins` élément doit être placé immédiatement avant la balise de fermeture `VersionOverrides` . Pour plus d’informations sur la compatibilité des compléments COM, consultez [la rubrique faire en sorte que votre complément Excel soit compatible avec un complément COM existant](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).

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
  <EquivalentAddins>
</VersionOverrides>
```

> [!NOTE]
> Si un complément déclare ses fonctions personnalisées comme étant compatibles XLL, la modification du manifeste ultérieurement pourrait entraîner la rupture du classeur d’un utilisateur, car il modifiera le format de fichier.

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>Comportement des fonctions personnalisées pour les fonctions compatibles XLL

Lors de l’ouverture d’une feuille de calcul qui contient des fonctions XLL pour lesquelles il existe également un complément équivalent, les fonctions de la XLL sont converties en fonctions personnalisées compatibles XLL. Lors du prochain enregistrement, les utilisateurs sont écrits dans le fichier dans un mode compatible afin qu’ils fonctionnent avec les fonctions personnalisées XLL et complément Excel (sur d’autres plateformes).

Le tableau suivant compare les fonctionnalités des fonctions définies par l’utilisateur XLL, des fonctions personnalisées de XLL et des fonctions personnalisées de complément Excel.

|         |Fonction XLL définie par l’utilisateur |Fonctions personnalisées compatibles XLL |Fonction personnalisée de complément Excel |
|---------|---------|---------|---------|
| Plateformes prises en charge | Windows | Windows, macOS, navigateur Web | Windows, macOS, navigateur Web |
| Formats de fichiers pris en charge | XLSX, XLSB, XLSM, XLS | XLSX, XLSB, XLSM | XLSX, XLSB, XLSM |
| Saisie semi-automatique de formule | Non | Oui | Oui |
| Diffusion en continu | Possible via xlfRTD et le rappel XLL. | Non | Oui |
| Localisation des fonctions | Non | Non. Le nom et l’ID doivent correspondre aux fonctions de la XLL existante. | Oui |
| Fonctions volatiles | Oui | Oui | Oui |
| Prise en charge du recalcul multi-thread | Oui | Oui | Oui |
| Comportement du calcul | Aucune interface utilisateur. Excel peut ne pas répondre pendant le calcul. | Les utilisateurs verront #BUSY ! jusqu’à ce qu’un résultat soit renvoyé. | Les utilisateurs verront #BUSY ! jusqu’à ce qu’un résultat soit renvoyé. |
| Ensembles de conditions requises | N/A | CustomFunctions 1,1 et versions ultérieures | CustomFunctions 1,1 et versions ultérieures |

## <a name="see-also"></a>Voir aussi

- [Faire en sorte que votre complément Excel soit compatible avec un complément COM existant](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
