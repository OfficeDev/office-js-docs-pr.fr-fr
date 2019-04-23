---
ms.date: 04/18/2019
description: Résoudre des problèmes courants dans les fonctions personnalisées d’Excel.
title: Résoudre des problèmes de fonctions personnalisées (préversion)
localization_priority: Priority
ms.openlocfilehash: cf54aa3b719b7893799df5d1c5206c6fb904be69
ms.sourcegitcommit: 44c61926d35809152cbd48f7b97feb694c7fa3de
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/22/2019
ms.locfileid: "31959103"
---
# <a name="troubleshoot-custom-functions"></a>Résoudre des problèmes de fonctions personnalisées

Dans le cadre du développement de fonctions personnalisées, vous pouvez rencontrer des erreurs dans le produit lors de la création et des tests de vos fonctions.

Pour résoudre des problèmes, vous pouvez [activer la journalisation du runtime pour capturer les erreurs](#enable-runtime-logging) et vous référer aux [messages d’erreur natifs d’Excel](#check-for-excel-error-messages). Recherchez également des erreurs courantes telles qu’une [vérification des certificats SSL](#my-add-in-wont-load-verify-certificates) incorrecte, l’[abandon de promesses non résolues](#ensure-promises-return) et l’oubli d’[associer votre fonctions](#my-functions-wont-load-associate-functions).

## <a name="enable-runtime-logging"></a>Activer la journalisation du runtime

Si vous testez votre complément dans Office sur Windows, vous devez [activer la journalisation du runtime](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in). La journalisation du runtime fournit des instructions `console.log` dans un fichier journal distinct que vous créez pour vous aider à découvrir des problèmes. Les instructions couvrent diverses erreurs, dont des erreurs liées au fichier manifeste XML de votre complément, aux conditions d’exécution ou à l’installation de vos fonctions personnalisées.  Pour plus d’informations sur la journalisation du runtime, voir [Utilisation de la journalisation du runtime pour déboguer votre complément](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).  

### <a name="check-for-excel-error-messages"></a>Rechercher les messages d’erreur Excel

Excel dispose d’un certain nombre de messages d’erreur intégrés qui sont renvoyés à une cellule en cas d’erreur de calcul. Les fonctions personnalisées utilisent uniquement les messages d’erreur suivants : `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A` et `#BUSY!`.

## <a name="common-issues"></a>Problèmes courants

### <a name="my-add-in-wont-load-verify-certificates"></a>Mon complément ne se charge pas : vérifiez les certificats

Si l’installation de votre complément échoue, vérifiez que les certificats SSL sont correctement configurés pour le serveur web hébergeant votre complément. Généralement, en cas de problème avec des certificats SSL, un message d’erreur dans Excel vous avertit que votre complément n’a pas pu être installé correctement. Pour plus d’informations, voir la rubrique relative à l’[ajout de certificats auto-signés en tant que certificats racine approuvés](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

### <a name="my-functions-wont-load-associate-functions"></a>Mon fonctions ne se chargent pas : associez les fonctions

Dans le fichier de script de vos fonctions personnalisées, vous devez associer chacune de celles-ci à son ID spécifié dans le [fichier de métadonnées JSON](custom-functions-json.md). Cette opération est effectuée à l’aide de la méthode `CustomFunctions.associate()`. Cette méthode est généralement appelée après chaque fonction ou à la fin du fichier de script. Si une fonction personnalisée n’est pas associée, elle ne fonctionne pas.

L’exemple suivant présente une fonction d’ajout, suivie du nom de la fonction `add` associé à l’id JSON correspondant `ADD`.

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

Pour plus d’informations sur ce processus, voir [Mappage des noms de fonction aux métadonnées JSON](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata).

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exception"></a>Impossible d’ouvrir les compléments d’hôte local : utiliser une exception de bouclage local

Si vous voyez le message d’erreur « Nous ne pouvons pas ouvrir ce complément à partir de l’hôte local », vous devez activer une exception de bouclage local. Pour plus d’informations sur la façon de procéder, voir [cet article du support Microsoft](https://support.microsoft.com/fr-FR/help/4490419/local-loopback-exemption-does-not-work).

### <a name="ensure-promises-return"></a>Veiller au renvoi de promesses

Quand Excel attend la fin de l’exécution d’une fonction personnalisée, il affiche #OCCUPÉ! dans la cellule. Si votre code de fonction personnalisée renvoie une promesse sans que celle-ci renvoie de résultat, Excel continue d’afficher #OCCUPÉ!. Vérifiez vos fonctions pour vous assurer que les promesses renvoient correctement un résultat à une cellule.

## <a name="reporting-feedback"></a>Formulation de commentaires

Si vous rencontrez des problèmes non abordés ici, faites-le nous savoir. Il existe deux méthodes pour signaler des problèmes.

### <a name="in-excel-on-windows-or-mac"></a>Dans Excel sur Windows ou Mac

Si vous utilisez Excel pour Windows ou Mac, vous pouvez adresser un commentaire à l’équipe chargée de l’extensibilité d’Office directement à partir d’Excel. Pour ce faire, sélectionnez **Fichier -> Commentaires -> Envoyer un smiley mécontent**. Envoyer un smiley mécontent fournira les journaux nécessaires pour comprendre le problème que vous rencontrez.

### <a name="in-github"></a>Dans Github

N’hésitez pas à signaler un problème rencontré via la fonctionnalité « Commentaires sur le contenu » accessible au bas de chaque page de documentation, ou en [déclarant un nouveau problème directement dans le référentiel de fonctions personnalisées](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## <a name="see-also"></a>Voir aussi

* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Fonctions personnalisées changelog](custom-functions-changelog.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
