---
ms.date: 06/09/2022
description: Résoudre les problèmes courants liés aux fonctions personnalisées Excel.
title: Résoudre des problèmes de fonctions personnalisées
ms.localizationpriority: medium
ms.openlocfilehash: 89d90b6ee94efac0230933313d2c16b5054dda61
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958558"
---
# <a name="troubleshoot-custom-functions"></a>Résoudre des problèmes de fonctions personnalisées

Dans le cadre du développement de fonctions personnalisées, vous pouvez rencontrer des erreurs dans le produit lors de la création et des tests de vos fonctions.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Pour résoudre des problèmes, vous pouvez [activer la journalisation du runtime pour capturer les erreurs](#enable-runtime-logging) et vous référer aux [messages d’erreur natifs d’Excel](#check-for-excel-error-messages). Recherchez également des erreurs courantes telles que l’[abandon de promesses non résolues](#ensure-promises-return).

## <a name="debugging-custom-functions"></a>Débogage de fonctions personnalisées

Pour déboguer des compléments de fonctions personnalisées qui utilisent un runtime partagé, consultez [Configurer votre complément Office pour utiliser un runtime JavaScript partagé : Déboguer](../develop/configure-your-add-in-to-use-a-shared-runtime.md#debug).

Pour déboguer des compléments de fonctions personnalisées qui n’utilisent pas de runtime partagé, consultez [Débogage des fonctions personnalisées](custom-functions-debugging.md).

## <a name="enable-runtime-logging"></a>Activer la journalisation du runtime

Si vous testez votre complément dans Office sur Windows, vous devez [activer la journalisation du runtime](../testing/runtime-logging.md). La journalisation du runtime fournit des instructions `console.log` dans un fichier journal distinct que vous créez pour vous aider à découvrir des problèmes. Les instructions couvrent diverses erreurs, dont des erreurs liées au fichier manifeste XML de votre complément, aux conditions d’exécution ou à l’installation de vos fonctions personnalisées. Pour plus d’informations sur la journalisation du runtime, voir [Déboguer votre complément à l’aide de la journalisation du runtime](../testing/runtime-logging.md).

### <a name="check-for-excel-error-messages"></a>Rechercher les messages d’erreur Excel

Excel dispose d’un certain nombre de messages d’erreur intégrés qui sont renvoyés à une cellule en cas d’erreur de calcul. Les fonctions personnalisées utilisent uniquement les messages d’erreur suivants : `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A` et `#BUSY!`.

En règle générale, ces erreurs correspondent aux erreurs que vous devez déjà connaître dans Excel. Il existe quelques exceptions spécifiques aux fonctions personnalisées et répertoriées ici :

- Une erreur `#NAME` indique généralement un problème d’inscription de vos fonctions.
- Une erreur `#N/A` peut également indiquer que l’exécution d’une fonction, bien qu’enregistrée, a échoué. Cet échec est généralement dû à une commande `CustomFunctions.associate` manquante.
- Une erreur `#VALUE` indique généralement une erreur dans le fichier de script des fonctions.
- Une erreur `#REF!` peut indiquer que le nom de votre fonction est identique au nom d’une fonction de complément déjà présent.

## <a name="clear-the-office-cache"></a>Vider le cache Office

Les informations relatives aux fonctions personnalisées sont mises en cache par Office. Lorsque vous développez et rechargez de manière répétée un complément avec des fonctions personnalisées, il peut arriver que modifications n’apparaissent pas. Pour y remédier, videz le cache Office. Pour plus d’informations, voir [Vider le cache Office](../testing/clear-cache.md).

## <a name="common-problems-and-solutions"></a>Problèmes et solutions courants

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exemption"></a>Impossible d’ouvrir le complément à partir de localhost : utiliser une exemption de bouclage local

Si vous voyez l’erreur « Nous ne pouvons pas ouvrir ce complément à partir de localhost », vous devez activer une exemption de bouclage locale. Pour plus d’informations sur la façon de procéder, voir [cet article du support Microsoft](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).

### <a name="runtime-logging-reports-typeerror-network-request-failed-on-excel-on-windows"></a>Rapports de journalisation d’exécution « TypeError: Network request failed » (« TypeError : échec de la requête réseau ») dans Excel sur Windows

Si le message d’erreur « TypeError: Network request failed » (« TypeError : échec de la requête réseau ») figure dans votre [journal d’exécution](custom-functions-troubleshooting.md#enable-runtime-logging) lorsque vous appelez votre serveur localhost, vous devez activer une exception de bouclage locale. Pour plus d’informations sur la façon de procéder, voir la *deuxième option* décrite dans [cet article du support Microsoft](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).

### <a name="ensure-promises-return"></a>Veiller au renvoi de promesses

Quand Excel attend la fin de l’exécution d’une fonction personnalisée, il affiche #OCCUPÉ! dans la cellule. Si votre code de fonction personnalisée renvoie une promesse sans que celle-ci renvoie de résultat, Excel continue d’afficher `#BUSY!`. Vérifiez vos fonctions pour vous assurer que les promesses renvoient correctement un résultat à une cellule.

### <a name="error-the-dev-server-is-already-running-on-port-3000"></a>Erreur : le serveur de développement est déjà en cours d’exécution sur le port 3000

Lorsque vous exécutez `npm start`, une erreur indiquant que le serveur de développement est déjà en cours d’exécution sur le port 3000 (ou le port utilisé par votre complément) peut s’afficher. Vous pouvez arrêter le serveur de développement en exécutant `npm stop` ou en fermant la fenêtre Node.js. Dans certains cas, l’arrêt de l’exécution du serveur de développement peut prendre quelques minutes.

### <a name="my-functions-wont-load-associate-functions"></a>Mon fonctions ne se chargent pas : associer les fonctions

Dans les cas où votre JSON n’a pas été inscrit et que vous avez créé vos propres métadonnées JSON, il se peut qu’un erreur `#VALUE!` s’affiche ou que vous receviez une notification indiquant que votre complément ne peut pas être chargé. Cela signifie généralement que vous devez associer chacune de vos fonctions personnalisées à sa propriété `id` spécifiée dans le [fichier de métadonnées JSON](custom-functions-json.md). Pour ce faire, utilisez la `CustomFunctions.associate()` fonction. En règle générale, cet appel de fonction est effectué après chaque fonction ou à la fin du fichier de script. Si une fonction personnalisée n’est pas associée, elle ne fonctionne pas.

L’exemple suivant présente une fonction d’ajout, suivie du nom de la fonction `add` associé à l’id JSON correspondant `ADD`.

```js
/**
 * Add two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

Pour plus d’informations sur ce processus, consultez [Association de noms de fonctions avec des métadonnées JSON](../excel/custom-functions-json.md#associating-function-names-with-json-metadata).

## <a name="known-issues"></a>Problèmes connus

Les problèmes connus sont suivis et signalés dans le [référentiel GitHub Excel Custom Functions](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## <a name="reporting-feedback"></a>Formulation de commentaires

Si vous rencontrez des problèmes non abordés ici, faites-le nous savoir. Il existe deux méthodes pour signaler des problèmes.

### <a name="in-excel-on-windows-or-mac"></a>Dans Excel sur Windows ou Mac

Si vous utilisez Excel sur Windows ou Mac, vous pouvez adresser un commentaire à l’équipe chargée de l’extensibilité d’Office directement à partir d’Excel. Pour ce faire, sélectionnez **Fichier -> Commentaires -> Envoyer un smiley mécontent**. Envoyer un smiley mécontent fournira les journaux nécessaires pour comprendre le problème que vous rencontrez.

### <a name="in-github"></a>Dans Github

N’hésitez pas à signaler un problème rencontré via la fonctionnalité « Commentaires sur le contenu » accessible au bas de chaque page de documentation, ou en [déclarant un nouveau problème directement dans le référentiel de fonctions personnalisées](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## <a name="next-steps"></a>Étapes suivantes

Découvrez comment [rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](make-custom-functions-compatible-with-xll-udf.md).

## <a name="see-also"></a>Voir aussi

- [Générer automatiquement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md)
- [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
- [Débogage des fonctions personnalisées](custom-functions-debugging.md)