---
title: Valider et résoudre des problèmes avec votre manifeste
description: Utiliser ces méthodes pour valider le manifeste des compléments Office.
ms.date: 11/02/2018
ms.openlocfilehash: 710a06108206675a6c4fe523137f12a5d12f1da4
ms.sourcegitcommit: c6723a31b48945ca4c466ba016a3dfc7b6267f5c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/03/2018
ms.locfileid: "25942243"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a>Valider et résoudre des problèmes avec votre manifeste

Utiliser les méthodes suivantes pour valider et résoudre les problèmes rencontrés dans votre manifeste pour compléments Office : 

- [Validation du manifeste à l’aide du validateur de complément Office](#validate-your-manifest-with-the-office-add-in-validator)   
- [Validation de votre manifeste par rapport au schéma XML](#validate-your-manifest-against-the-xml-schema)
- [Validation de votre manifeste avec le générateur Yeoman pour les compléments Office](#validate-your-manifest-with-the-yeoman-generator-for-office-add-ins)
- [Utilisation de la journalisation runtime pour déboguer le manifeste de votre complément](#use-runtime-logging-to-debug-your-add-in-manifest)


## <a name="validate-your-manifest-with-the-office-add-in-validator"></a>Validation du manifeste à l’aide du validateur de complément Office

Pour vous assurer que le fichier manifeste qui décrit votre complément Office est correct et complet, vérifiez-le à l’aide du [validateur de complément Office](https://github.com/OfficeDev/office-addin-validator).

### <a name="to-use-the-office-add-in-validator-to-validate-your-manifest"></a>Pour utiliser le validateur de complément Office afin de valider votre manifeste

1. Installez [Node.js](https://nodejs.org/download/). 

2. Ouvrez une invite de commandes/un terminal en tant qu’administrateur, puis installez le validateur de complément Office et ses dépendances de façon globale à l’aide de la commande suivante :

    ```bash
    npm install -g office-addin-validator
    ```
    
    > [!NOTE]
    > Si Yo Office est déjà installé, effectuez une mise à niveau vers la dernière version ; le validateur sera installé en tant que dépendance.

3. Exécutez la commande suivante pour valider votre manifeste. Remplacez MANIFEST.XML par le chemin d’accès au fichier XML de manifeste.

    ```bash
    validate-office-addin MANIFEST.XML
    ```

## <a name="validate-your-manifest-against-the-xml-schema"></a>Validation de votre manifeste par rapport au schéma XML

Assurez-vous que le fichier manifeste suit le schéma approprié, y compris les espaces de noms pour les éléments que vous utilisez. Si vous avez copié des éléments à partir d’autres exemples de manifestes, vérifiez par deux fois que vous avez également **inclus les espaces de noms appropriés**. Vous pouvez valider un manifeste par rapport aux fichiers de [définition de schéma XML (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas). Pour ce faire, vous pouvez utiliser un outil de validation de schéma XML. 



### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a>Pour utiliser un outil de validation de schéma XML à ligne de commande pour valider votre manifeste

1.  Installez [tar](https://www.gnu.org/software/tar/) et [libxml](http://xmlsoft.org/FAQ.html), si vous ne l’avez pas déjà fait.

2.  Exécutez la commande suivante. Remplacez `XSD_FILE` par le chemin d’accès au fichier XSD manifeste et `XML_FILE` par le chemin d’accès au fichier XML manifeste.
    
    ```bash
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a>Validation de votre manifeste avec le générateur Yeoman pour les compléments Office

Si vous avez créé votre complément Office à l’aide du [générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office), vérifiez que le fichier manifeste suit le schéma correct en exécutant la commande suivante dans le répertoire racine de votre projet :

```bash
npm run validate
```

![Gif animé qui montre le validateur Yo Office exécuté sur la ligne de commande et les résultats générés indiquant « Validation Passed » (validation réussie)](../images/yo-office-validator.gif)

> [!NOTE]
> Pour accéder à cette fonctionnalité, votre projet de complément doit être créé à l’aide du [générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) (version 1.1.17 ou ultérieure).

## <a name="use-runtime-logging-to-debug-your-add-in"></a>Utilisation de la journalisation runtime pour déboguer votre complément 

Vous pouvez utiliser la journalisation runtime pour déboguer le manifeste de votre complément ainsi que plusieurs erreurs d’installation. Cette fonctionnalité peut vous aider à identifier et à résoudre les problèmes avec votre manifeste qui ne sont pas détectés par la validation de schéma XSD, comme une incompatibilité entre les ID de ressources. La journalisation runtime est particulièrement utile pour déboguer des compléments qui implémentent des commandes de complément et des fonctions personnalisées Excel.   

> [!NOTE]
> La fonctionnalité de journalisation runtime est actuellement disponible pour Office 2016 pour ordinateur de bureau.

### <a name="to-turn-on-runtime-logging"></a>Pour activer la journalisation runtime

> [!IMPORTANT]
> La journalisation runtime réduit les performances. Activez-la uniquement lorsque vous avez besoin de déboguer des problèmes avec votre manifeste de complément.

Pour activer la journalisation runtime, procédez comme suit :

1. Vérifiez que vous exécutez la version Bureau d’Office 2016 **16.0.7019** ou une version ultérieure. 

2. Ajoutez la clé de registre `RuntimeLogging` sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`. 

    > [!NOTE]
    > Si la clé (dossier) `Developer` n’existe pas sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, procédez comme suit pour la créer : 
    > 1. Cliquez avec le bouton droit de votre souris sur la clé (dossier) **WEF**, puis sélectionnez **Nouveau** > **Clé**.
    > 2. Nommez la nouvelle clé **Développeur**.

3. Définissez la valeur par défaut de la clé pour le chemin d’accès complet du fichier dans lequel écrire le journal. Pour obtenir un exemple, reportez-vous à [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip). 

    > [!NOTE]
    > Le répertoire dans lequel le fichier journal sera écrit doit déjà exister et vous devez disposer des autorisations d’écriture correspondantes. 
 
L’image suivante indique à quoi doit ressembler le registre. Pour désactiver la fonctionnalité, supprimez la clé de registre `RuntimeLogging`. 

![Capture d’écran de l’Éditeur du Registre avec la clé de registre RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)


### <a name="to-troubleshoot-issues-with-your-manifest"></a>Résolution des problèmes avec votre manifeste

Pour utiliser la journalisation runtime pour résoudre les problèmes de chargement d’un complément, procédez comme suit :
 
1. [Chargez une version test de votre complément](sideload-office-add-ins-for-testing.md). 

    > [!NOTE]
    > Nous vous recommandons de charger uniquement une version test du complément que vous testez pour réduire le nombre de messages dans le fichier journal.

2. Si rien ne se produit et que votre complément n’apparaît pas (et ne s’affiche pas dans la boîte de dialogue des compléments), ouvrez le fichier journal.

3. Recherchez le fichier journal pour l’ID de votre complément, que vous définissez dans votre manifeste. Dans le fichier journal, cet ID est intitulé `SolutionId`. 

Dans l’exemple suivant, le fichier journal identifie un contrôle qui pointe vers un fichier de ressources qui n’existe pas. Pour cet exemple, la correction consistera à corriger la faute de frappe dans le manifeste ou à ajouter la ressource manquante.

![Capture d’écran d’un fichier journal avec une entrée qui spécifie un ID de ressource qui est introuvable](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a>Problèmes connus avec la journalisation runtime

Vous pouvez afficher des messages dans le fichier journal qui sont source de confusion ou classés de façon incorrecte. Par exemple :

- Le message `Medium Current host not in add-in's host list` suivi de `Unexpected Parsed manifest targeting different host` est classé incorrectement en tant qu’erreur.

- Si vous voyez le message `Unexpected Add-in is missing required manifest fields DisplayName` et qu’il ne contient pas de SolutionId, l’erreur n’est probablement pas liée au complément que vous déboguez. 

- Tous les messages `Monitorable` sont des erreurs attendues du point de vue du système. Parfois, ils indiquent un problème avec votre manifeste, comme un élément mal orthographié qui a été ignoré, mais n’a pas provoqué l’échec du manifeste. 

## <a name="clear-the-office-cache"></a>Vider le cache Office

Si les modifications apportées au manifeste, par exemple aux noms de fichier des icônes de bouton dans le ruban ou au texte des commandes de complément, ne semblent pas appliquées, essayez de vider le cache Office de votre ordinateur. 

#### <a name="for-windows"></a>Pour Windows :
Supprimez le contenu du dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

#### <a name="for-mac"></a>Pour Mac :
Supprimez le contenu du dossier `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.

#### <a name="for-ios"></a>Pour iOS :
Appelez `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également choisir de réinstaller Office.

## <a name="see-also"></a>Voir aussi

- [Manifeste XML des compléments Office](../develop/add-in-manifests.md)
- [Chargement de la version test des compléments Office](sideload-office-add-ins-for-testing.md)
- [Débogage des compléments Office](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
