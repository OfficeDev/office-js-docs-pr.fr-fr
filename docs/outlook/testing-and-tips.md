---
title: Déployer et installer des compléments Outlook à des fins de test
description: Créez un fichier manifeste, déployez le fichier IU de complément, installez le complément dans votre boîte aux lettres, puis testez-le.
ms.date: 11/06/2019
localization_priority: Priority
ms.openlocfilehash: 521199a87282b58c3bf10553886174e8be26cacf
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166074"
---
# <a name="deploy-and-install-outlook-add-ins-for-testing"></a>Déployer et installer des compléments Outlook à des fins de test

Dans le cadre du processus de développement d’un complément Outlook, vous devrez déployer et installer de façon itérative le complément à des fins de test, ce qui implique les étapes suivantes :

1. Création d’un fichier manifeste qui décrit le complément.
1. Déploiement du ou des fichiers de l’interface utilisateur du complément sur un serveur web.
1. Installation du complément dans votre boîte aux lettres.
1. Test du complément, mise en œuvre des modifications appropriées dans l’interface utilisateur ou dans les fichiers manifeste, et répétition des étapes 2 et 3 pour tester les modifications.

> [!NOTE]
> [Les volets personnalisés sont déconseillés](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) afin de vous assurer que vous utilisez [un point d’extension de complément pris en charge](outlook-add-ins-overview.md#extension-points).

## <a name="create-a-manifest-file-for-the-add-in"></a>Création d’un fichier manifeste pour le complément

Chaque complément est décrit par un manifeste XML, un document qui fournit au serveur des informations sur le complément, décrit le complément pour l’utilisateur et identifie l’emplacement du fichier HTML de l’interface utilisateur du complément. Vous pouvez stocker le manifeste dans un dossier local ou sur un serveur, à condition que le complément soit accessible par le serveur Exchange de la boîte aux lettres avec laquelle vous procédez aux tests. Nous partons du principe que vous stockez votre manifeste dans un dossier local. Pour plus d’informations sur la création d’un fichier manifeste, voir [Manifestes des compléments Outlook](manifests.md).

## <a name="deploy-an-add-in-to-a-web-server"></a>Déploiement d’un complément sur un serveur web

Vous pouvez utiliser du code HTML et JavaScript pour créer le complément. Les fichiers source obtenus sont stockés sur un serveur web accessible par le biais du serveur Exchange qui héberge le complément. Après le déploiement initial des fichiers source pour le complément, vous pouvez mettre à jour l’interface utilisateur et le comportement du complément en remplaçant les fichiers HTML ou JavaScript stocké sur le serveur web par une nouvelle version du fichier HTML.

## <a name="install-the-add-in"></a>Installer le complément

Après la préparation du fichier manifeste du complément et le déploiement de son interface utilisateur sur un serveur web accessible, vous pouvez charger une version test du complément pour une boîte aux lettres sur un serveur Exchange à l’aide d’un client Outlook ou installer le complément en exécutant des cmdlets Windows PowerShell à distance.

### <a name="sideload-the-add-in"></a>Charger une version test du complément

Vous pouvez installer un complément si votre boîte aux lettres est sur Exchange Online, Exchange 2013 ou une version ultérieure. Les compléments de chargement de version test nécessitent au minimum le rôle **Mes compléments personnalisés** pour votre serveur Exchange. Pour tester votre complément ou installer des compléments en général en spécifiant une URL ou un nom de fichier pour le manifeste de complément, vous devez demander à votre administrateur Exchange de vous octroyer les autorisations nécessaires.

L’administrateur Exchange peut exécuter la cmdlet PowerShell suivante pour affecter les autorisations nécessaires à un seul utilisateur. Dans cet exemple, `wendyri` est l’alias de messagerie de l’utilisateur.

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

Selon les besoins, l’administrateur peut exécuter la cmdlet suivante pour affecter des autorisations nécessaires similaires à plusieurs utilisateurs :

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

Pour plus d’informations sur le rôle « Mes compléments personnalisés », consultez la rubrique relative au [rôle « Mes compléments personnalisés »](/exchange/my-custom-apps-role-exchange-2013-help).

L’utilisation d’Office 365 ou de Visual Studio pour développer des compléments vous amène à endosser le rôle d’administrateur d’organisation, ce qui vous permet d’installer des compléments par fichier ou par URL dans le Centre d’administration Exchange ou via des cmdlets PowerShell.

### <a name="install-an-add-in-by-using-remote-powershell"></a>Installation d’un complément à l’aide de PowerShell à distance

Après avoir créé une session Windows PowerShell à distance sur votre serveur Exchange, vous pouvez installer un complément Outlook en utilisant la cmdlet `New-App` avec la commande PowerShell suivante.

```powershell
New-App -URL:"http://<fully-qualified URL">
```

L’URL complète est l’emplacement du fichier de manifeste de complément que vous avez préparé pour votre complément.

Vous pouvez utiliser les cmdlets supplémentaires suivantes pour gérer les compléments pour une boîte aux lettres :

-  `Get-App` : répertorie les compléments activés pour une boîte aux lettres.
-  `Set-App` : active ou désactive un complément sur une boîte aux lettres.
-  `Remove-App` : supprime un complément précédemment installé à partir d’un serveur Exchange.

## <a name="client-versions"></a>Versions client

Le choix des versions du client Outlook à tester dépend de vos besoins en matière de développement.

- Si vous développez un complément pour une utilisation privée ou uniquement pour les membres de votre organisation, il est important de tester les versions d’Outlook que votre entreprise utilise. Gardez à l’esprit que certains utilisateurs peuvent utiliser Outlook sur le web. Par conséquent, vous devez également tester les versions des navigateurs standard utilisés au sein de votre entreprise.

- Si vous développez un complément pour [AppSource](https://appsource.microsoft.com), vous devez tester les versions requises tel que spécifié dans les [stratégies de validation d’AppSource 4.12.1](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably). Cela inclut notamment :
    - la dernière et avant-dernière version d’Outlook sur Windows ;
    - la dernière version d’Outlook sur Mac ;
    - la dernière version d’Outlook sur iOS et Android (si votre complément [prend en charge le facteur de forme pour mobile](add-mobile-support.md)) ;
    - les versions de navigateur spécifiées dans la stratégie de validation d’AppSource 4.12.1.

> [!NOTE]
> Si votre complément ne prend pas en charge l’un des clients ci-dessus car il demande [un ensemble de conditions requises d’API](apis.md) que le client ne prend pas en charge, ce client est supprimé de la liste des clients requis.

## <a name="see-also"></a>Voir aussi

- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)
