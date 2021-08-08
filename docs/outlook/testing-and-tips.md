---
title: Déployer et installer des compléments Outlook à des fins de test
description: Créez un fichier manifeste, déployez le fichier IU de complément, installez le complément dans votre boîte aux lettres, puis testez-le.
ms.date: 07/08/2021
localization_priority: Priority
ms.openlocfilehash: 0fe7aa8d24b4da14a14480aaf07ef588cd8a243a
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773089"
---
# <a name="deploy-and-install-outlook-add-ins-for-testing"></a>Déployer et installer des compléments Outlook à des fins de test

Dans le cadre du processus de développement d’un complément Outlook, vous serez probablement amené à déployer et à installer le complément de manière itérative à des fins de test, ce qui implique les étapes suivantes.

1. Création d’un fichier manifeste qui décrit le complément.
1. Déploiement du ou des fichiers de l’interface utilisateur du complément sur un serveur web.
1. Installation du complément dans votre boîte aux lettres.
1. Test du complément, mise en œuvre des modifications appropriées dans l’interface utilisateur ou dans les fichiers manifeste, et répétition des étapes 2 et 3 pour tester les modifications.

> [!NOTE]
> [Les volets personnalisés sont déconseillés](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) afin de vous assurer que vous utilisez [un point d’extension de complément pris en charge](outlook-add-ins-overview.md#extension-points).

## <a name="create-a-manifest-file-for-the-add-in"></a>Création d’un fichier manifeste pour le complément

Chaque complément est décrit par un manifeste XML, un document qui fournit au serveur des informations sur le complément, décrit le complément pour l’utilisateur et identifie l’emplacement du fichier HTML de l’interface utilisateur du complément. Vous pouvez stocker le manifeste dans un dossier ou un serveur local, tant que l’emplacement est accessible par le serveur Exchange de la boîte aux lettres que vous testez. Nous partons du principe que vous stockez votre manifeste dans un dossier local. Pour plus d’informations sur la création d’un fichier manifeste, voir [Manifestes des compléments Outlook](manifests.md).

## <a name="deploy-an-add-in-to-a-web-server"></a>Déploiement d’un complément sur un serveur web

Vous pouvez utiliser du code HTML et JavaScript pour créer le complément. Les fichiers source obtenus sont stockés sur un serveur web accessible par le biais du serveur Exchange qui héberge le complément. Après le déploiement initial des fichiers source pour le complément, vous pouvez mettre à jour l’interface utilisateur et le comportement du complément en remplaçant les fichiers HTML ou JavaScript stocké sur le serveur web par une nouvelle version du fichier HTML.

## <a name="install-the-add-in"></a>Installer le complément

Après la préparation du fichier manifeste du complément et le déploiement de son interface utilisateur sur un serveur web accessible, vous pouvez charger une version test du complément pour une boîte aux lettres sur un serveur Exchange à l’aide d’un client Outlook ou installer le complément en exécutant des cmdlets Windows PowerShell à distance.

### <a name="sideload-the-add-in"></a>Charger une version test du complément

Vous pouvez installer un complément si votre boîte aux lettres est sur Exchange Online, Exchange 2013 ou une version ultérieure. Les compléments de chargement de version test nécessitent au minimum le rôle **Mes compléments personnalisés** pour votre serveur Exchange. Pour tester votre complément ou installer des compléments en général en spécifiant une URL ou un nom de fichier pour le manifeste de complément, vous devez demander à votre administrateur Exchange de vous octroyer les autorisations nécessaires.

L’administrateur Exchange peut exécuter la cmdlet PowerShell suivante pour affecter les autorisations nécessaires à un seul utilisateur. Dans cet exemple, `wendyri` est l’alias de messagerie de l’utilisateur.

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

Si nécessaire, l’administrateur peut exécuter la cmdlet suivante pour attribuer à plusieurs utilisateurs des autorisations nécessaires similaires.

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

Pour plus d’informations sur le rôle « Mes compléments personnalisés », consultez la rubrique relative au [rôle « Mes compléments personnalisés »](/exchange/my-custom-apps-role-exchange-2013-help).

Utiliser Microsoft 365 ou Visual Studio pour développer des add-ins vous attribue le rôle d'administrateur de l'organisation, ce qui vous permet d'installer des add-ins par fichier ou URL dans l'EAC, ou par cmdlets Powershell.

### <a name="install-an-add-in-by-using-remote-powershell"></a>Installation d’un complément à l’aide de PowerShell à distance

Après avoir créé une session Windows PowerShell à distance sur votre serveur Exchange, vous pouvez installer un complément Outlook en utilisant la cmdlet `New-App` avec la commande PowerShell suivante.

```powershell
New-App -URL:"http://<fully-qualified URL">
```

L’URL complète est l’emplacement du fichier de manifeste de complément que vous avez préparé pour votre complément.

Utilisez les applets de commande PowerShell supplémentaires suivantes pour gérer les compléments d’une boîte aux lettres.

- `Get-App` : répertorie les compléments activés pour une boîte aux lettres.
- `Set-App` : active ou désactive un complément sur une boîte aux lettres.
- `Remove-App` : supprime un complément précédemment installé à partir d’un serveur Exchange.

## <a name="client-versions"></a>Versions client

Le choix des versions du client Outlook à tester dépend de vos besoins en matière de développement.

- Si vous développez un complément pour une utilisation privée ou uniquement pour les membres de votre organisation, il est important de tester les versions d’Outlook utilisées par votre entreprise. N’oubliez pas que certains utilisateurs peuvent utiliser Outlook sur le web. Il est donc également important de tester les versions de navigateur standard de votre entreprise.

- Si vous développez un complément à répertorier dans [AppSource](https://appsource.microsoft.com), vous devez tester les versions requises comme spécifié dans les [stratégies de certification de la place de marché commerciale 1120.3](/legal/marketplace/certification-policies#11203-functionality). Cela inclut les opérations suivantes :
  - la dernière et avant-dernière version d’Outlook sur Windows ;
  - la dernière version d’Outlook sur Mac ;
  - la dernière version d’Outlook sur iOS et Android (si votre complément [prend en charge le facteur de forme pour mobile](add-mobile-support.md)) ;
  - Les versions de navigateur spécifiées dans la stratégie de validation de la Place de marché commerciale 1120.3.

> [!NOTE]
> Si votre complément ne prend pas en charge l’un des clients ci-dessus car il demande [un ensemble de conditions requises d’API](apis.md) que le client ne prend pas en charge, ce client est supprimé de la liste des clients requis.

## <a name="outlook-on-the-web-and-exchange-server-versions"></a>Outlook sur le web et les versions du serveur Exchange

Les utilisateurs de comptes consommateurs et de Microsoft 365 voient la version moderne de l'interface utilisateur lorsqu'ils accèdent à Outlook sur le web et ne voient plus la version classique qui a été dépréciée. Toutefois, les serveurs Exchange sur site continuent de prendre en charge le protocole Outlook classique sur le web. Par conséquent, pendant le processus de validation, votre soumission peut recevoir un avertissement indiquant que le module complémentaire n'est pas compatible avec Outlook classique sur le web. Dans ce cas, vous devriez envisager de tester votre add-in dans un environnement d'échange sur site. Cet avertissement ne bloquera pas votre soumission à AppSource, mais vos clients risquent de vivre une expérience non optimale s'ils utilisent Outlook sur le web dans un environnement Exchange sur site.

Pour atténuer ce problème, nous vous recommandons de tester votre module d'extension dans Outlook sur le web, connecté à votre propre environnement Exchange privé sur site. Pour plus d'informations, voir les conseils sur la façon d´[Établir un environnement d'essai pour Exchange 2016 ou Exchange 2019](/Exchange/plan-and-deploy/plan-and-deploy?view=exchserver-2019&preserve-view=true#establish-an-exchange-2016-or-exchange-2019-test-environment) et comment gérer [Outlook on the web in Exchange Server](/exchange/clients/outlook-on-the-web/outlook-on-the-web?view=exchserver-2019&preserve-view=true).

Autrement, vous pouvez également choisir de payer et d'utiliser un service qui héberge et gère sur son site des serveurs Exchange. Voici quelques options:

- [Rackspace](https://www.rackspace.com/email-hosting/exchange-server)
- [Hostway](https://hostway.com/microsoft-exchange/)

En outre, si vous ne souhaitez pas que vos add-ins soient disponibles pour les utilisateurs qui sont connectés à l'échange sur site, vous pouvez définir le [exigences fixées](../reference/requirement-sets/outlook-api-requirement-sets.md#exchange-server-support) dans le manifeste de l'add-in à 1,6 ou plus. Ces add-ins ne seront pas testés ou validés sur l'interface utilisateur classique de Outlook on the web.

## <a name="see-also"></a>Voir aussi

- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)
