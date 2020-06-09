---
title: Déboguez votre complément avec la journalisation runtime
description: Découvrez l’utilisation de la journalisation runtime pour déboguer votre complément.
ms.date: 12/31/2019
localization_priority: Normal
ms.openlocfilehash: e97f989ff4791da678fbecbf8bceed3670749e01
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611146"
---
# <a name="debug-your-add-in-with-runtime-logging"></a>Déboguez votre complément avec la journalisation runtime

Vous pouvez utiliser la journalisation runtime pour déboguer le manifeste de votre complément ainsi que plusieurs erreurs d’installation. Cette fonctionnalité peut vous aider à identifier et à résoudre les problèmes avec votre manifeste qui ne sont pas détectés par la validation de schéma XSD, comme une incompatibilité entre les ID de ressources. La journalisation runtime est particulièrement utile pour déboguer des compléments qui implémentent des commandes de complément et des fonctions personnalisées Excel.   

> [!NOTE]
> La fonctionnalité de journalisation runtime est actuellement disponible pour Office 2016 pour ordinateur de bureau.

> [!IMPORTANT]
> La journalisation runtime affecte les performances. Activez-la uniquement lorsque vous avez besoin de déboguer des problèmes avec votre manifeste de complément.

## <a name="use-runtime-logging-from-the-command-line"></a>Utiliser la journalisation de l’exécution à partir de la ligne de commande

L’activation de la journalisation de l’exécution à partir de la ligne de commande est le moyen le plus rapide d’utiliser cet outil de journalisation. Celles-ci utilisent npx, fourni par défaut dans le cadre de npm@5.2.0 +. Si vous disposez d’une version antérieure de [npm](https://www.npmjs.com/), essayez les instructions [Journalisation de l’exécution sur Windows](#runtime-logging-on-windows) ou [Journalisation de l’exécution sur Mac](#runtime-logging-on-mac), ou [install npx](https://www.npmjs.com/package/npx).

- Pour activer la journalisation de l’exécution :
    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable
    ```
- Pour activer la journalisation de l’exécution uniquement pour un fichier spécifique, utilisez la même commande avec un nom de fichier :

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable [filename.txt]
    ```

- Pour désactiver la journalisation de l’exécution :

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --disable
    ```

- Pour indiquer si la journalisation de l’exécution est activée :

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log
    ```

- Pour afficher l’aide au sein de la ligne de commande pour la journalisation de l’exécution :

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --help
    ```

## <a name="runtime-logging-on-windows"></a>Journalisation de l’exécution sur Windows

1. Vérifiez que vous exécutez la version Bureau d’Office 2016 **16.0.7019** ou une version ultérieure. 

2. Ajoutez la clé de registre `RuntimeLogging` sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`. 

    > [!NOTE]
    > Si la clé (dossier) `Developer` n’existe pas sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, procédez comme suit pour la créer : 
    > 1. Cliquez avec le bouton droit de votre souris sur la clé (dossier) **WEF**, puis sélectionnez **Nouveau** > **Clé**.
    > 2. Nommez la nouvelle clé **Développeur**.

3. Définissez la valeur par défaut de la clé **RuntimeLogging** pour le chemin d’accès complet du fichier dans lequel écrire le journal. Pour obtenir un exemple, voir [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip). 

    > [!NOTE]
    > Le répertoire dans lequel le fichier journal sera écrit doit déjà exister et vous devez disposer des autorisations d’écriture correspondantes. 
 
L’image suivante indique à quoi doit ressembler le registre. Pour désactiver la fonctionnalité, supprimez la clé de registre `RuntimeLogging`. 

![Capture d’écran de l’Éditeur du Registre avec la clé de registre RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)

## <a name="runtime-logging-on-mac"></a>Journalisation de l’exécution sur Mac

1. Vérifiez que vous exécutez la version Bureau d’Office 2016 **16.27** (19071500) ou une version ultérieure.

2. Ouvrez **Terminal** et configurez une préférence de journalisation de l’exécution à l’aide de la commande `defaults` :
    
    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    `<bundle id>` identifie l’hôte pour lequel activer la journalisation de l’exécution. `<file_name>` est le nom du fichier texte dans lequel le journal sera écrit.

    Configurez `<bundle id>` à l’une des valeurs suivantes pour activer la journalisation de l’exécution pour l’hôte correspondant :

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

L’exemple suivant montre comment activer la journalisation de l’exécution pour Word, puis ouvrir le fichier journal :

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE] 
> Vous devrez redémarrer Office après l’exécution de la commande `defaults` pour activer la journalisation de l’exécution.

Pour désactiver la journalisation de l’exécution, utilisez la commande `defaults delete` :

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

L’exemple suivant désactive la journalisation de l’exécution pour Word :

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

## <a name="use-runtime-logging-to-troubleshoot-issues-with-your-manifest"></a>Utilisez la journalisation runtime pour résoudre les problèmes avec votre manifeste

Pour utiliser la journalisation runtime pour résoudre les problèmes de chargement d’un complément, procédez comme suit :
 
1. [Chargez une version test de votre complément](sideload-office-add-ins-for-testing.md). 

    > [!NOTE]
    > Nous vous recommandons de charger uniquement une version test du complément que vous testez pour réduire le nombre de messages dans le fichier journal.

2. Si rien ne se produit et que votre complément n’apparaît pas (et ne s’affiche pas dans la boîte de dialogue des compléments), ouvrez le fichier journal.

3. Recherchez le fichier journal pour l’ID de votre complément, que vous définissez dans votre manifeste. Dans le fichier journal, cet ID est intitulé `SolutionId`. 

Dans l’exemple suivant, le fichier journal identifie un contrôle qui pointe vers un fichier de ressources qui n’existe pas. Pour cet exemple, la correction consistera à corriger la faute de frappe dans le manifeste ou à ajouter la ressource manquante.

![Capture d’écran d’un fichier journal avec une entrée qui spécifie un ID de ressource qui est introuvable](http://i.imgur.com/f8bouLA.png) 

## <a name="known-issues-with-runtime-logging"></a>Problèmes connus avec la journalisation runtime

Vous pouvez afficher des messages dans le fichier journal qui sont source de confusion ou classés de façon incorrecte. Par exemple :

- Le message `Medium Current host not in add-in's host list` suivi de `Unexpected Parsed manifest targeting different host` est classé incorrectement en tant qu’erreur.

- Si vous voyez le message `Unexpected Add-in is missing required manifest fields    DisplayName` et qu’il ne contient pas de SolutionId, l’erreur n’est probablement pas liée au complément que vous déboguez. 

- Tous les messages `Monitorable` sont des erreurs attendues du point de vue du système. Parfois, ils indiquent un problème avec votre manifeste, comme un élément mal orthographié qui a été ignoré, mais n’a pas provoqué l’échec du manifeste. 

## <a name="see-also"></a>Voir aussi

- [Manifeste XML des compléments Office](../develop/add-in-manifests.md)
- [Valider le manifeste d’un complément Office](troubleshoot-manifest.md)
- [Vider le cache Office](clear-cache.md)
- [Chargement de la version test des compléments Office](sideload-office-add-ins-for-testing.md)
- [Débogage des compléments Office](debug-add-ins-using-f12-developer-tools-on-windows-10.md)