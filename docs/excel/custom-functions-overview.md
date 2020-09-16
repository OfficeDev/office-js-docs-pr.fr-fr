---
ms.date: 08/13/2020
description: Créez une fonction personnalisée Excel pour votre Complément Office.
title: Créer des fonctions personnalisées dans Excel
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 731e8d99a36cfef7d125838c67efcdd7a77b4bb1
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819559"
---
# <a name="create-custom-functions-in-excel"></a>Créer des fonctions personnalisées dans Excel

Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément. Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

L’image animée suivante montre votre classeur appelant une fonction que vous avez créée avec JavaScript ou Typescript. Dans cet exemple, la fonction personnalisée `=MYFUNCTION.SPHEREVOLUME` calcule le volume d’une sphère.

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

Le code suivant définit la fonction personnalisée `=MYFUNCTION.SPHEREVOLUME`.

```js
/**
 * Returns the volume of a sphere.
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
```

> [!NOTE]
> La section [problèmes connus](#known-issues)plus loin dans cet article indique les limitations en cours de fonctions personnalisées.

## <a name="how-a-custom-function-is-defined-in-code"></a>Comment une fonction personnalisée est définie dans le code

Si vous utilisez le [générateur de Yo Office](https://github.com/OfficeDev/generator-office) pour créer un projet de complément de fonctions personnalisées Excel, il crée des fichiers qui contrôlent totalement vos fonctions, et volet des tâches. Nous allons vous concentrer sur les fichiers importants pour les fonctions personnalisées :

| File | Format de fichier | Description |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>ou<br/>**./src/functions/functions.ts** | JavaScript<br/>ou<br/>TypeScript | Contient le code qui définit les fonctions personnalisées. |
| **./src/functions/functions.html** | HTML | Fournit une référence&lt;script&gt; au fichier JavaScript qui définit les fonctions personnalisées. |
| **./manifest.xml** | XML | Indique l’emplacement de plusieurs fichiers utilisés par votre fonction personnalisée, tels que les fonctions personnalisées JavaScript, JSON et HTML. Il répertorie également les emplacements des fichiers du volet Office, des fichiers de commandes et indique le runtime que vos fonctions personnalisées doivent utiliser. |

### <a name="script-file"></a>Fichier de script

Le fichier de script (**./src/functions/functions.js** ou **./src/functions/functions.ts**) contient le code qui définit des fonctions personnalisées et des commentaires qui définissent la fonction.

Le code suivant définit la fonction personnalisée `add`. Les commentaires du code sont utilisés pour générer un fichier de métadonnées JSON décrivant la fonction personnalisée pour Excel. Le commentaire obligatoire `@customfunction` est déclaré en premier, pour indiquer qu’il s’agit d’une fonction personnalisée. Deux paramètres sont ensuite déclarés, `first` et `second`, suivis de leurs propriétés de `description` . Enfin, une description `returns` est fournie. Pour plus d’informations sur les commentaires requis pour votre fonction personnalisée, voir [Créer des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md).

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number.
 * @param second Second number.
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}
```

### <a name="manifest-file"></a>Fichier manifeste

Le fichier manifeste XML pour un complément qui définit des fonctions personnalisées (**./manifest.xml** dans le projet que le générateur de bureau Yo crée) effectue plusieurs opérations :

- Définit l’espace de noms pour vos fonctions personnalisées. Un espace de noms s’ajoute à vos fonctions personnalisées pour aider les clients à identifier vos fonctions dans le cadre de votre complément.
- Utilise les éléments `<ExtensionPoint>` et `<Resources>` qui sont propres à un manifeste de fonctions personnalisées. Ces éléments contiennent les informations relatives aux emplacements des fichiers JavaScript, JSON et HTML.
- Spécifie le runtime à utiliser pour votre fonction personnalisée. Nous vous recommandons de toujours utiliser une exécution partagée, sauf si vous avez un besoin spécifique d’autre runtime, car un runtime partagé autorise le partage de données entre les fonctions et le volet Office. Notez que l’utilisation d’un runtime partagé signifie que votre complément utilise Internet Explorer 11, et non Microsoft Edge.

Si vous utilisez le générateur Yo Office pour créer des fichiers, nous vous recommandons d’ajuster votre manifeste pour utiliser un runtime partagé, car il ne s’agit pas de la valeur par défaut pour ces fichiers. Pour modifier votre manifeste, suivez les instructions dans [Configurer votre complément Excel pour utiliser un runtime JavaScript partagé](configure-your-add-in-to-use-a-shared-runtime.md).

Pour afficher un manifeste de travail complet à partir d’un exemple de complément, consultez [ce référentiel GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a>Co-édition

Excel sur le web et Windows connecté à un abonnement Microsoft 365 vous permettent de co-éditer dans Excel. Si votre classeur utilise une fonction personnalisée, votre collègue coauteur est invité à charger le complément de la fonction personnalisée. Quand vous avez tous les deux chargé le complément, la fonction personnalisée partage les résultats via la co-édition.

Pour plus d’informations sur la co-création, voir [À propos de la co-création dans Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).

## <a name="known-issues"></a>Problèmes connus

Consulter les problèmes connus sur notre[repo GitHub Fonctions Excel Personnalisées](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## <a name="next-steps"></a>Étapes suivantes

Vous voulez essayer les fonctions personnalisées ? Consultez la documentation sur le [démarrage rapide de fonction personnalisée](../quickstarts/excel-custom-functions-quickstart.md) ou le [didacticiel sur les fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md).

Un autre moyen simple d’essayer des fonctions personnalisées consiste à utiliser [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), un complément qui vous permet d’expérimenter des fonctions personnalisées directement dans Excel. Vous pouvez essayer de créer votre propre fonction personnalisée ou utiliser les exemples fournis.

## <a name="see-also"></a>Voir aussi 
* [Configuration requise de fonctions personnalisées](custom-functions-requirement-sets.md)
* [Instructions d’attribution de noms](custom-functions-naming.md)
* [Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](make-custom-functions-compatible-with-xll-udf.md)
