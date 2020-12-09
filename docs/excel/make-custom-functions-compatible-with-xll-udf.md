---
title: Étendre des fonctions personnalisées avec des fonctions XLL définies par l’utilisateur
description: Activer la compatibilité avec les fonctions Excel XLL définies par l’utilisateur qui offrent une fonctionnalité équivalente à vos fonctions personnalisées
ms.date: 08/13/2020
localization_priority: Normal
ms.openlocfilehash: c34dcf5ef546fa0f337b2cbd11cca7d5e25e2de3
ms.sourcegitcommit: fecad2afa7938d7178456c11ba52b558224813b4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/09/2020
ms.locfileid: "49603777"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a>Étendre des fonctions personnalisées avec des fonctions XLL définies par l’utilisateur

Si vous avez des XLL Excel existantes, vous pouvez créer des fonctions personnalisées équivalentes dans un complément Excel afin d’étendre les fonctionnalités de votre solution à d’autres plateformes, comme Online ou sur Mac. Toutefois, les compléments Excel ne disposent pas de toutes les fonctionnalités disponibles dans les XLL. En fonction de la fonctionnalité utilisée par votre solution, le XLL peut offrir une meilleure expérience que les fonctions personnalisées de complément Excel dans Excel sur Windows.

> [!NOTE]
> Le complément COM et la compatibilité UDF XLL sont pris en charge par les plateformes suivantes, lorsqu’ils sont connectés à un abonnement Microsoft 365 :
> - Excel sur le web
> - Excel sur Windows (version 1904 ou ultérieure)
> - Excel sur Mac (version 13,329 ou ultérieure)
>
> Pour utiliser le complément COM et la compatibilité des FDU XLL dans Excel sur le Web, connectez-vous à l’aide de votre abonnement Microsoft 365 ou d’un [compte Microsoft](https://account.microsoft.com/account). Si vous ne disposez pas déjà d’un abonnement Microsoft 365, vous pouvez obtenir gratuitement un abonnement Microsoft 365 renouvelable 90 jours en joignant le [programme de développement microsoft 365](https://developer.microsoft.com/office/dev-program).

## <a name="specify-equivalent-xll-in-the-manifest"></a>Spécifier le XLL équivalent dans le manifeste

Pour activer la compatibilité avec un XLL existant, identifiez le XLL équivalent dans le manifeste de votre complément Excel. Excel utilise ensuite les fonctions de la XLL au lieu de vos fonctions personnalisées de complément Excel lors de l’exécution de Windows.

Pour définir le XLL équivalent pour vos fonctions personnalisées, spécifiez l’élément `FileName` XLL. Lorsque l’utilisateur ouvre un classeur avec des fonctions à partir de la XLL, Excel convertit les fonctions en fonctions compatibles. Le classeur utilise ensuite le XLL lorsqu’il est ouvert dans Excel sur Windows et utilise des fonctions personnalisées à partir de votre complément Excel lorsqu’il est ouvert en ligne ou sur un Mac.

L’exemple suivant montre comment spécifier un complément COM et un XLL comme équivalent. Vous devez souvent spécifier les deux. Cet exemple montre des éléments à la fois dans le contexte. Ils sont identifiés par leur `ProgId` et `FileName` respectivement. L' `EquivalentAddins` élément doit être placé immédiatement avant la `VersionOverrides` balise de fermeture. Pour plus d’informations sur la compatibilité des compléments COM, consultez [la rubrique faire en sorte que votre complément Excel soit compatible avec un complément COM existant](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).

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
> Si un complément déclare ses fonctions personnalisées comme étant compatibles XLL, la modification du manifeste ultérieurement pourrait entraîner la rupture du classeur d’un utilisateur, car il modifiera le format de fichier.

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>Comportement des fonctions personnalisées pour les fonctions compatibles XLL

Les fonctions XLL d’un complément sont converties en fonctions personnalisées compatibles XLL lorsqu’une feuille de calcul est ouverte et qu’un complément équivalent est disponible. Lors du prochain enregistrement, les fonctions XLL sont écrites dans le fichier dans un mode compatible de sorte qu’elles fonctionnent avec les fonctions personnalisées XLL et complément Excel (sur d’autres plateformes).

Le tableau suivant compare les fonctionnalités des fonctions définies par l’utilisateur XLL, des fonctions personnalisées de XLL et des fonctions personnalisées de complément Excel.

|         |Fonction XLL définie par l’utilisateur |Fonctions personnalisées compatibles XLL |Fonction personnalisée de complément Excel |
|---------|---------|---------|---------|
| **Plateformes prises en charge** | Windows | Windows, macOS, navigateur Web | Windows, macOS, navigateur Web |
| **Formats de fichiers pris en charge** | XLSX, XLSB, XLSM, XLS | XLSX, XLSB, XLSM | XLSX, XLSB, XLSM |
| **Saisie semi-automatique de formule** | Non | Oui | Oui |
| **Diffusion en continu** | Possible via xlfRTD et le rappel XLL. | Oui | Oui |
| **Localisation des fonctions** | Non | Non. Le nom et l’ID doivent correspondre aux fonctions de la XLL existante. | Oui |
| **Fonctions volatiles** | Oui | Oui | Oui |
| **Prise en charge du recalcul multi-thread** | Oui | Oui | Oui |
| **Comportement du calcul** | Aucune interface utilisateur. Excel peut ne pas répondre pendant le calcul. | Les utilisateurs verront #BUSY ! jusqu’à ce qu’un résultat soit renvoyé. | Les utilisateurs verront #BUSY ! jusqu’à ce qu’un résultat soit renvoyé. |
| **Ensembles de conditions requises** | S/O | CustomFunctions 1,1 et versions ultérieures | CustomFunctions 1,1 et versions ultérieures |

## <a name="see-also"></a>Voir aussi

- [Faire en sorte que votre complément Excel soit compatible avec un complément COM existant](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
