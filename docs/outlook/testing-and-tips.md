---
title: Déployer et installer des compléments Outlook à des fins de test
description: Créez un fichier manifeste, déployez le fichier IU de complément, installez le complément dans votre boîte aux lettres, puis testez-le.
ms.date: 05/20/2020
localization_priority: Priority
ms.openlocfilehash: 97841f7c8112b42cee2927f238b31fe985b2e101
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093860"
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

Each add-in is described by an XML manifest, a document that gives the server information about the add-in, provides descriptive information about the add-in for the user, and identifies the location of the add-in UI HTML file. You can store the manifest in a local folder or server, as long as the location is accessible by the Exchange server of the mailbox that you are testing with. We'll assume that you store your manifest in a local folder. For information about how to create a manifest file, see [Outlook add-in manifests](manifests.md).

## <a name="deploy-an-add-in-to-a-web-server"></a>Déploiement d’un complément sur un serveur web

You can use HTML and JavaScript to create the add-in. The resulting source files are stored on a web server that can be accessed by the Exchange server that hosts the add-in. After initially deploying the source files for the add-in, you can update the add-in UI and behavior by replacing the HTML files or JavaScript files stored on the web server with a new version of the HTML file.

## <a name="install-the-add-in"></a>Installer le complément

Après la préparation du fichier manifeste du complément et le déploiement de son interface utilisateur sur un serveur web accessible, vous pouvez charger une version test du complément pour une boîte aux lettres sur un serveur Exchange à l’aide d’un client Outlook ou installer le complément en exécutant des cmdlets Windows PowerShell à distance.

### <a name="sideload-the-add-in"></a>Charger une version test du complément

You can install an add-in if your mailbox is on Exchange Online, Exchange 2013 or a later release. Sideloading add-ins requires at minimum the **My Custom Apps** role for your Exchange Server. In order to test your add-in, or install add-ins in general by specifying a URL or file name for the add-in manifest, you should request your Exchange administrator to provide the necessary permissions.

The Exchange administrator can run the following PowerShell cmdlet to assign a single user the necessary permissions. In this example, `wendyri` is the user's email alias.

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

Selon les besoins, l’administrateur peut exécuter la cmdlet suivante pour affecter des autorisations nécessaires similaires à plusieurs utilisateurs :

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

Pour plus d’informations sur le rôle « Mes compléments personnalisés », consultez la rubrique relative au [rôle « Mes compléments personnalisés »](/exchange/my-custom-apps-role-exchange-2013-help).

Utiliser Microsoft 365 ou Visual Studio pour développer des add-ins vous attribue le rôle d'administrateur de l'organisation, ce qui vous permet d'installer des add-ins par fichier ou URL dans l'EAC, ou par cmdlets Powershell.

### <a name="install-an-add-in-by-using-remote-powershell"></a>Installation d’un complément à l’aide de PowerShell à distance

Après avoir créé une session Windows PowerShell à distance sur votre serveur Exchange, vous pouvez installer un complément Outlook en utilisant la cmdlet `New-App` avec la commande PowerShell suivante.

```powershell
New-App -URL:"http://<fully-qualified URL">
```

L’URL complète est l’emplacement du fichier de manifeste de complément que vous avez préparé pour votre complément.

Vous pouvez utiliser les cmdlets supplémentaires suivantes pour gérer les compléments pour une boîte aux lettres :

- `Get-App` : répertorie les compléments activés pour une boîte aux lettres.
- `Set-App` : active ou désactive un complément sur une boîte aux lettres.
- `Remove-App` : supprime un complément précédemment installé à partir d’un serveur Exchange.

## <a name="client-versions"></a>Versions client

Le choix des versions du client Outlook à tester dépend de vos besoins en matière de développement.

- If you are developing an add-in for private use, or only for members of your organization, then it is important to test the versions of Outlook that your company uses. Keep in mind that some users may use Outlook on the web, so testing your company's standard browser versions is also important.

- If you are developing an add-in to list in [AppSource](https://appsource.microsoft.com), you must test the required versions as specified in the [Commercial marketplace certification policies 1120.3](/legal/marketplace/certification-policies#11203-functionality). This includes:
  - la dernière et avant-dernière version d’Outlook sur Windows ;
  - la dernière version d’Outlook sur Mac ;
  - la dernière version d’Outlook sur iOS et Android (si votre complément [prend en charge le facteur de forme pour mobile](add-mobile-support.md)) ;
  - Les versions de navigateur spécifiées dans la stratégie de validation de la Place de marché commerciale 1120.3.

> [!NOTE]
> Si votre complément ne prend pas en charge l’un des clients ci-dessus car il demande [un ensemble de conditions requises d’API](apis.md) que le client ne prend pas en charge, ce client est supprimé de la liste des clients requis.

## <a name="outlook-on-the-web-and-exchange-server-versions"></a>Outlook sur le web et les versions du serveur Exchange

Les utilisateurs de comptes consommateurs et de Microsoft 365 voient la version moderne de l'interface utilisateur lorsqu'ils accèdent à Outlook sur le web et ne voient plus la version classique qui a été dépréciée. Toutefois, les serveurs Exchange sur site continuent de prendre en charge le protocole Outlook classique sur le web. Par conséquent, pendant le processus de validation, votre soumission peut recevoir un avertissement indiquant que le module complémentaire n'est pas compatible avec Outlook classique sur le web. Dans ce cas, vous devriez envisager de tester votre add-in dans un environnement d'échange sur site. Cet avertissement ne bloquera pas votre soumission à AppSource, mais vos clients risquent de vivre une expérience non optimale s'ils utilisent Outlook sur le web dans un environnement Exchange sur site.

Pour atténuer ce problème, nous vous recommandons de tester votre module d'extension dans Outlook sur le web, connecté à votre propre environnement Exchange privé sur site. Pour plus d'informations, voir les conseils sur la façon d´[Établir un environnement d'essai pour Exchange 2016 ou Exchange 2019](/Exchange/plan-and-deploy/plan-and-deploy?view=exchserver-2019#establish-an-exchange-2016-or-exchange-2019-test-environment) et comment gérer [Outlook on the web in Exchange Server](/exchange/clients/outlook-on-the-web/outlook-on-the-web?view=exchserver-2019).

Vous pouvez également choisir de payer et d'utiliser un service qui héberge et gère des serveurs Exchange sur place. Il existe plusieurs options :

- [Rackspace](https://www.rackspace.com/email-hosting/exchange-server)
- [Hostway](https://hostway.com/products-services-2/hosted-microsoft-exchange/)

En outre, si vous ne souhaitez pas que vos add-ins soient disponibles pour les utilisateurs qui sont connectés à l'échange sur site, vous pouvez définir le [exigences fixées](../reference/requirement-sets/outlook-api-requirement-sets.md#exchange-server-support) dans le manifeste de l'add-in à 1,6 ou plus. Ces add-ins ne seront pas testés ou validés sur l'interface utilisateur classique de Outlook on the web.

## <a name="see-also"></a>Voir aussi

- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)
