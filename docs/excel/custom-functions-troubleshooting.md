---
ms.date: 06/21/2019
description: Résoudre des problèmes courants dans les fonctions personnalisées d’Excel.
title: Résoudre des problèmes de fonctions personnalisées
localization_priority: Priority
ms.openlocfilehash: f42e9e6ac3e2d868f90ab4f5129684308750b8e2
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128282"
---
# <a name="troubleshoot-custom-functions"></a>Résoudre des problèmes de fonctions personnalisées

Dans le cadre du développement de fonctions personnalisées, vous pouvez rencontrer des erreurs dans le produit lors de la création et des tests de vos fonctions.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Pour résoudre des problèmes, vous pouvez [activer la journalisation du runtime pour capturer les erreurs](#enable-runtime-logging) et vous référer aux [messages d’erreur natifs d’Excel](#check-for-excel-error-messages). Recherchez également des erreurs courantes telles que l’[abandon de promesses non résolues](#ensure-promises-return) et l’oubli d’[associer vos fonctions](#my-functions-wont-load-associate-functions).

## <a name="enable-runtime-logging"></a>Activer la journalisation du runtime

Si vous testez votre complément dans Office sur Windows, vous devez [activer la journalisation du runtime](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in). La journalisation du runtime fournit des instructions `console.log` dans un fichier journal distinct que vous créez pour vous aider à découvrir des problèmes. Les instructions couvrent diverses erreurs, dont des erreurs liées au fichier manifeste XML de votre complément, aux conditions d’exécution ou à l’installation de vos fonctions personnalisées.  Pour plus d’informations sur la journalisation du runtime, voir [Utilisation de la journalisation du runtime pour déboguer votre complément](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).  

### <a name="check-for-excel-error-messages"></a>Rechercher les messages d’erreur Excel

Excel dispose d’un certain nombre de messages d’erreur intégrés qui sont renvoyés à une cellule en cas d’erreur de calcul. Les fonctions personnalisées utilisent uniquement les messages d’erreur suivants : `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A` et `#BUSY!`.

En règle générale, ces erreurs correspondent aux erreurs que vous devez déjà connaître dans Excel. Il existe quelques exceptions spécifiques aux fonctions personnalisées et répertoriées ici :

- Une erreur `#NAME` indique généralement un problème d’inscription de vos fonctions.
- Une erreur `#VALUE` indique généralement une erreur dans le fichier de script des fonctions.
- Une erreur `#N/A` peut également indiquer que l’exécution d’une fonction, bien qu’enregistrée, a échoué. Cet échec est généralement dû à une commande `CustomFunctions.associate` manquante.
- Une erreur `#REF!` peut indiquer que le nom de votre fonction est identique au nom d’une fonction de complément déjà présent.

## <a name="clear-the-office-cache"></a>Vider le cache Office

Les informations relatives aux fonctions personnalisées sont mises en cache par Office. Lorsque vous développez et rechargez de manière répétée un complément avec des fonctions personnalisées, il peut arriver que modifications n’apparaissent pas. Pour y remédier, videz le cache Office. Pour plus d’informations, consultez la section « Vider le cache Office » de l’article [Valider et résoudre des problèmes avec votre manifeste](https://docs.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest?branch=master#clear-the-office-cache).

## <a name="common-issues"></a>Problèmes courants

### <a name="my-functions-wont-load-associate-functions"></a>Mon fonctions ne se chargent pas : associez les fonctions

Dans le fichier de script de vos fonctions personnalisées, vous devez associer chacune de celles-ci à son ID spécifié dans le [fichier de métadonnées JSON](custom-functions-json.md). Cette opération est effectuée à l’aide de la méthode `CustomFunctions.associate()`. Cette méthode est généralement appelée après chaque fonction ou à la fin du fichier de script. Si une fonction personnalisée n’est pas associée, elle ne fonctionne pas.

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

Pour plus d’informations sur ce processus, voir [Mappage des noms de fonction aux métadonnées JSON](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata).

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exception"></a>Impossible d’ouvrir les compléments d’hôte local : utiliser une exception de bouclage local

Si vous voyez le message d’erreur « Nous ne pouvons pas ouvrir ce complément à partir de l’hôte local », vous devez activer une exception de bouclage local. Pour plus d’informations sur la façon de procéder, voir [cet article du support Microsoft](https://support.microsoft.com/fr-FR/help/4490419/local-loopback-exemption-does-not-work).

### <a name="runtime-logging-reports-typeerror-network-request-failed-on-excel-on-windows"></a>Rapports de journalisation d’exécution « TypeError: Network request failed » (« TypeError : échec de la requête réseau ») dans Excel sur Windows

Si le message d’erreur « TypeError: Network request failed » (« TypeError : échec de la requête réseau ») figure dans votre [journal d’exécution](custom-functions-troubleshooting.md#enable-runtime-logging) lorsque vous appelez votre serveur localhost, vous devez activer une exception de bouclage locale. Pour plus d’informations sur la façon de procéder, voir la *deuxième option* décrite dans [cet article du support Microsoft](https://support.microsoft.com/fr-FR/help/4490419/local-loopback-exemption-does-not-work).

### <a name="ensure-promises-return"></a>Veiller au renvoi de promesses

Quand Excel attend la fin de l’exécution d’une fonction personnalisée, il affiche #OCCUPÉ! dans la cellule. Si votre code de fonction personnalisée renvoie une promesse sans que celle-ci renvoie de résultat, Excel continue d’afficher #OCCUPÉ!. Vérifiez vos fonctions pour vous assurer que les promesses renvoient correctement un résultat à une cellule.

### <a name="error-the-dev-server-is-already-running-on-port-3000"></a>Erreur : le serveur de développement est déjà en cours d’exécution sur le port 3000

Lorsque vous exécutez `npm start`, une erreur indiquant que le serveur de développement est déjà en cours d’exécution sur le port 3000 (ou le port utilisé par votre complément) peut s’afficher. Vous pouvez arrêter le serveur de développement en exécutant `npm stop` ou en fermant la fenêtre Node.js. Notez que l’arrêt du serveur de développement peut prendre quelques minutes.

## <a name="reporting-feedback"></a>Formulation de commentaires

Si vous rencontrez des problèmes non abordés ici, faites-le nous savoir. Il existe deux méthodes pour signaler des problèmes.

### <a name="in-excel-on-windows-or-mac"></a>Dans Excel sur Windows ou Mac

Si vous utilisez Excel sur Windows ou Mac, vous pouvez adresser un commentaire à l’équipe chargée de l’extensibilité d’Office directement à partir d’Excel. Pour ce faire, sélectionnez **Fichier -> Commentaires -> Envoyer un smiley mécontent**. Envoyer un smiley mécontent fournira les journaux nécessaires pour comprendre le problème que vous rencontrez.

### <a name="in-github"></a>Dans Github

N’hésitez pas à signaler un problème rencontré via la fonctionnalité « Commentaires sur le contenu » accessible au bas de chaque page de documentation, ou en [déclarant un nouveau problème directement dans le référentiel de fonctions personnalisées](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## <a name="next-steps"></a>Étapes suivantes
Découvrez comment [déboguer vos fonctions personnalisées](custom-functions-debugging.md).

## <a name="see-also"></a>Voir aussi

* [Génération automatique de métadonnées de fonctions personnalisées](custom-functions-json-autogeneration.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](make-custom-functions-compatible-with-xll-udf.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
