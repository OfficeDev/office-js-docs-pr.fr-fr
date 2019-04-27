---
title: Faire en sorte que vos fonctions personnalisées soient compatibles avec les fonctions XLL définies par l'utilisateur
description: Activer la compatibilité avec les fonctions Excel XLL définies par l'utilisateur qui offrent une fonctionnalité équivalente à vos fonctions personnalisées
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 09914e040c1721dd8b9e91952e5814e7a6b914e5
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356865"
---
# <a name="make-your-custom-functions-compatible-with-xll-user-defined-functions"></a>Faire en sorte que vos fonctions personnalisées soient compatibles avec les fonctions XLL définies par l'utilisateur

Si vous avez des XLL Excel existantes, vous pouvez créer des fonctions personnalisées équivalentes dans un complément Office pour étendre les fonctionnalités de votre solution à d'autres plateformes, comme Online ou macOS. Toutefois, les compléments Office ne disposent pas de toutes les fonctionnalités disponibles dans les XLL. En fonction de la fonctionnalité utilisée par votre solution, le XLL peut offrir une meilleure expérience que les fonctions personnalisées de complément Office sur Excel pour Windows.

Vous pouvez configurer votre complément Office de sorte que, lorsqu'un XLL équivalent est déjà installé sur l'ordinateur de l'utilisateur, Excel exécute le XLL à la place de vos fonctions personnalisées de complément Office. La XLL est appelée équivalente, car Excel effectuera une transition transparente entre les fonctions personnalisées XLL et complément Office en fonction de ce qui est installé sur Windows.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="specify-equivalent-xll-in-the-manifest"></a>Spécifier le XLL équivalent dans le manifeste

Pour activer la compatibilité avec un XLL existant, identifiez le XLL équivalent dans le manifeste de votre complément Office. Ensuite, Excel utilise les fonctions de la XLL au lieu des fonctions personnalisées de votre complément Office lors de l'exécution de Windows.

Pour définir le XLL équivalent pour vos fonctions personnalisées, spécifiez l' `FileName` élément XLL. Lorsque l'utilisateur ouvre un classeur avec des fonctions à partir de la XLL, Excel convertit les fonctions en fonctions compatibles. Le classeur utilise ensuite le XLL lorsqu'il est ouvert dans Excel sur Windows et utilise des fonctions personnalisées à partir de votre complément Office lorsqu'il est ouvert en ligne ou sur macOS.

L'exemple suivant montre comment spécifier un complément COM et un XLL comme équivalent. Souvent, vous spécifierez à la fois de manière à ce que cet exemple montre les deux dans le contexte. Ils sont identifiés par leur `ProgID` et `FileName` respectivement. Pour plus d'informations sur la compatibilité des compléments COM, consultez [la rubrique faire en sorte que votre complément Office soit compatible avec un complément COM existant](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).

```xml
<VersionOverrides>
...
<EquivalentAddins>
  <EquivalentAddin>
    <ProgID>ContosoCOMAddin</ProgID>
    <Type>COM</Type>
  </EquivalentAddin>

  <EquivalentAddin>
    <FileName>contosofunctions.xll</FileName>
    <Type>XLL</Type>
  </EquivalentAddin>
<EquivalentAddins>
...
</VersionOverrides>
```

> [!NOTE]
> Si un complément déclare ses fonctions personnalisées comme étant compatibles XLL, la modification du manifeste ultérieurement pourrait entraîner la rupture du classeur d'un utilisateur, car il modifiera le format de fichier.

## <a name="office-add-in-updates"></a>Mises à jour des compléments Office

Une fois que vous avez spécifié une XLL équivalente pour votre complément Office, Excel cesse de traiter les mises à jour pour votre complément Office. L'utilisateur doit désinstaller le XLL afin d'obtenir les dernières mises à jour pour le complément Office.

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>Comportement des fonctions personnalisées pour les fonctions compatibles XLL

Lors de l'ouverture d'une feuille de calcul qui contient des fonctions XLL pour lesquelles il existe également un complément équivalent, les fonctions de la XLL sont converties en fonctions personnalisées compatibles XLL. Lors du prochain enregistrement, les utilisateurs sont écrits dans le fichier dans un mode compatible afin qu'ils fonctionnent avec les fonctions personnalisées XLL et complément Office (sur d'autres plateformes).

Le tableau suivant compare les fonctionnalités des fonctions définies par l'utilisateur XLL, des fonctions personnalisées de XLL et des fonctions personnalisées de complément Office.

|         |Fonction XLL définie par l'utilisateur |Fonctions personnalisées compatibles XLL |Fonction personnalisée de complément Office |
|---------|---------|---------|---------|
| Plateformes prises en charge | Windows | Windows, macOS, Excel Online | Windows, macOS, Excel Online |
| Formats de fichiers pris en charge | XLSX, XLSB, XLSM, XLS | XLSX, XLSB, XLSM | XLSX, XLSB, XLSM |
| Saisie semi-automatique de formule | Non | Oui | Oui |
| Diffusion en continu | Possible via xlfRTD et le rappel XLL. | Oui | Oui |
| Localisation des fonctions | Non | Non. Le nom et l'ID doivent correspondre aux fonctions de la XLL existante. | Oui |
| Fonctions volatiles | Oui | Oui | Oui |
| Prise en charge du recalcul multi-thread | Oui | Oui | Oui |
| Comportement du calcul | Aucune interface utilisateur. Excel peut ne pas répondre pendant le calcul. | Les utilisateurs verront #BUSY! jusqu'à ce qu'un résultat soit renvoyé. | Les utilisateurs verront #BUSY! jusqu'à ce qu'un résultat soit renvoyé. |
| Ensembles de conditions requises | S/O | CustomFunctions 1,1 uniquement | CustomFunctions 1,1 et versions ultérieures |

## <a name="see-also"></a>Voir aussi

- [Faire en sorte que votre complément Office soit compatible avec un complément COM existant](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
- [Fonctions personnalisées changelog](custom-functions-changelog.md)
- [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)