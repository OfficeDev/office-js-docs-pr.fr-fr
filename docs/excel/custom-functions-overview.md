---
description: Créez une fonction personnalisée Excel pour votre Complément Office.
title: Créer des fonctions personnalisées dans Excel
ms.date: 08/04/2021
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 12740615215913b0201426f929dbcb803c866648
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422760"
---
# <a name="create-custom-functions-in-excel"></a>Créer des fonctions personnalisées dans Excel

Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément. Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

L’image animée suivante montre votre classeur appelant une fonction que vous avez créée avec JavaScript ou TypeScript. Dans cet exemple, la fonction personnalisée `=MYFUNCTION.SPHEREVOLUME` calcule le volume d’une sphère.

![Image animée montrant un utilisateur final insérant la fonction personnalisée MYFUNCTION.SPHEREVOLUME dans une cellule d’une feuille de calcul Excel.](../images/SphereVolumeNew.gif)

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

> [!TIP]
> Si votre complément de fonction personnalisée utilise un volet Office ou un bouton de ruban, en plus d’exécuter du code de fonction personnalisé, vous devez configurer un [runtime partagé](../testing/runtimes.md#shared-runtime). Pour plus d’informations, consultez [Configurer votre complément Office pour utiliser un runtime partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="how-a-custom-function-is-defined-in-code"></a>Comment une fonction personnalisée est définie dans le code

Si vous utilisez le [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md) pour créer un projet de complément de fonctions personnalisées Excel, il crée des fichiers qui contrôlent vos fonctions et votre volet Office. Nous nous concentrerons sur les fichiers importants pour les fonctions personnalisées.

| Fichier | Format de fichier | Description |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>ou<br/>**./src/functions/functions.ts** | JavaScript<br/>ou<br/>TypeScript | Contient le code qui définit les fonctions personnalisées. |
| **./src/functions/functions.html** | HTML | Fournit une référence&lt;script&gt; au fichier JavaScript qui définit les fonctions personnalisées. |
| **./manifest.xml** | XML | Indique l’emplacement de plusieurs fichiers utilisés par votre fonction personnalisée, tels que les fonctions personnalisées JavaScript, JSON et HTML. Il répertorie également les emplacements des fichiers du volet Office, des fichiers de commandes et indique le runtime que vos fonctions personnalisées doivent utiliser. |

### <a name="script-file"></a>Fichier de script

Le fichier de script (**./src/functions/functions.js** ou **./src/functions/functions.ts**) contient le code qui définit des fonctions personnalisées et des commentaires qui définissent la fonction.

Le code suivant définit la fonction personnalisée `add`. Les commentaires du code sont utilisés pour générer un fichier de métadonnées JSON décrivant la fonction personnalisée pour Excel. Le commentaire obligatoire `@customfunction` est déclaré en premier, pour indiquer qu’il s’agit d’une fonction personnalisée. Deux paramètres sont ensuite déclarés, `first` et `second`, suivis de leurs propriétés de `description` . Enfin, une description `returns` est fournie. Pour plus d’informations sur les commentaires requis pour votre fonction personnalisée, voir [Générer automatiquement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md).

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

Le fichier manifeste XML d’un complément qui définit des fonctions personnalisées (**./manifest.xml** dans le projet que le générateur [Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md) crée) effectue plusieurs opérations.

- Définit l’espace de noms pour vos fonctions personnalisées. Un espace de noms s’ajoute à vos fonctions personnalisées pour aider les clients à identifier vos fonctions dans le cadre de votre complément.
- Utilise les éléments **\<ExtensionPoint\>** et **\<Resources\>** qui sont propres à un manifeste de fonctions personnalisées. Ces éléments contiennent les informations relatives aux emplacements des fichiers JavaScript, JSON et HTML.
- Spécifie le runtime à utiliser pour votre fonction personnalisée. Nous vous recommandons de toujours utiliser une exécution partagée, sauf si vous avez un besoin spécifique d’autre runtime, car un runtime partagé autorise le partage de données entre les fonctions et le volet Office.

Si vous utilisez le [générateur Yeoman pour compléments Office pour créer des](../develop/yeoman-generator-overview.md) fichiers, nous vous recommandons d’ajuster votre manifeste pour utiliser un runtime partagé, car il ne s’agit pas de la valeur par défaut de ces fichiers. Pour modifier votre manifeste, suivez les instructions de [configuration de votre complément Excel pour utiliser un runtime partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

Pour afficher un manifeste de travail complet à partir d’un exemple de complément, consultez le manifeste dans le [l’un de nos exemples de dépôts Github de complément Office](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/excel-shared-runtime-global-state/manifest.xml).

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a>Co-édition

Excel sur le web et sur Windows connecté à un abonnement Microsoft 365 permettent aux utilisateurs finaux de co-éditer dans Excel. Si le classeur d’un utilisateur final utilise une fonction personnalisée, le collègue de co-création de cet utilisateur final est invité à charger le complément de fonctions personnalisées correspondant. Une fois que les deux utilisateurs ont chargé le complément, la fonction personnalisée partage les résultats via la co-édition.

Pour plus d’informations sur la co-création, voir [À propos de la co-création dans Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).

## <a name="next-steps"></a>Étapes suivantes

Vous voulez essayer des fonctions personnalisées ? Si ce n’est déjà fait, consultez le [démarrage rapide des fonctions personnalisées ](../quickstarts/excel-custom-functions-quickstart.md)simples ou le didacticiel plus détaillé [fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md).

Un autre moyen simple d’essayer des fonctions personnalisées consiste à utiliser [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), un complément qui vous permet d’expérimenter des fonctions personnalisées directement dans Excel. Vous pouvez essayer de créer votre propre fonction personnalisée ou utiliser les exemples fournis.

## <a name="see-also"></a>Voir aussi

- [Découvrez le programme pour les développeurs Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
- [Ensembles de besoins de fonctions personnalisées](/javascript/api/requirement-sets/excel/custom-functions-requirement-sets)
- [Règles de noms des fonctions personnalisées](custom-functions-naming.md)
- [Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](make-custom-functions-compatible-with-xll-udf.md)
- [Configurer votre complément Office pour utiliser un runtime partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Runtimes dans les compléments Office](../testing/runtimes.md)
