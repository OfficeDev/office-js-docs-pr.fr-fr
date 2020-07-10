---
title: Configuration requise pour exécuter des compléments Office
description: Découvrez la configuration requise du client et du serveur pour qu’un utilisateur final doive exécuter des compléments Office.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: fa01decddcc7cc59945ad92912fabab90cc505f7
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093482"
---
# <a name="requirements-for-running-office-add-ins"></a>Configuration requise pour exécuter des compléments Office

Cet article décrit la configuration logicielle et matérielle requise pour l’exécution des compléments Office.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

Pour savoir de manière détaillée quelle version d’Office prend en charge les compléments Office, consultez la page relative à la [disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md).

## <a name="server-requirements"></a>Exigences en matière de serveur

Pour pouvoir installer et exécuter des Complément Office, vous devez d’abord déployer les fichiers manifeste et de pages web pour l’interface utilisateur et le code de votre complément sur les emplacements de serveur appropriés.

Pour tous les types de complément (compléments de contenu, Outlook et volet Office, et les commandes de compléments), vous devez déployer les fichiers de pages web de votre complément sur un serveur web ou un service d’hébergement web, tel que [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> lorsque vous développez et déboguez un complément dans Visual Studio, Visual Studio déploie et exécute les fichiers de page web de votre complément localement avec IIS Express et ne nécessite aucun serveur web supplémentaires.

Pour les compléments du volet Office et de contenu, dans les applications hôtes Office prises en charge (Excel, PowerPoint, Project ou Word), vous avez également besoin d’un [catalogue d’applications](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) sur SharePoint pour télécharger les fichiers manifeste XML des compléments.

Pour tester et exécuter un complément Outlook, le compte de messagerie Outlook de l’utilisateur doit résider sur Exchange 2013 ou une version ultérieure, disponible via Microsoft 365, Exchange Online ou via une installation locale. L’utilisateur ou l’administrateur installe les fichiers manifeste pour les compléments Outlook sur ce serveur.

> [!NOTE]
> Les comptes de messagerie POP et IMAP dans Outlook ne prennent pas en charge les compléments Office.

## <a name="client-requirements-windows-desktop-and-tablet"></a>Exigences en matière de client : ordinateur de bureau et tablette Windows

Le logiciel suivant est requis pour développer un Complément Office pour les clients Office ou les clients web pris en charge qui s’exécutent sur un ordinateur de bureau, un ordinateur portable ou une tablette Windows :


- Pour les ordinateurs de bureau Windows x86 et x64 et les tablettes telles que Surface Pro :
    - La version 32 bits ou 64 bits d’Office 2013 ou une version ultérieure s’exécutant sur Windows 7 ou une version ultérieure.
    - Excel 2013, Outlook 2013, PowerPoint 2013, Project Professional 2013, Project 2013 SP1, Word 2013, or a later version of the Office client, if you are testing or running an Office Add-in specifically for one of these Office desktop clients. Office desktop clients can be installed on premises or via Click-to-Run on the client computer.

  Si vous disposez d’un abonnement Microsoft 365 valide et que vous n’avez pas accès au client Office, vous pouvez [Télécharger et installer la dernière version d’Office](https://support.office.com/article/download-and-install-or-reinstall-office-365-or-office-2019-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658).

- Internet Explorer 11 ou Microsoft Edge (selon les versions de Windows et Office) doit être installé, sans être nécessairement le navigateur par défaut. Pour prendre en charge les compléments Office, le client Office servant d’hôte utilise des composants de navigateur faisant partie d’Internet Explorer 11 ou de Microsoft Edge. Pour plus d’informations, voir [Navigateurs utilisés par les compléments Office](browsers-used-by-office-web-add-ins.md).

  > [!NOTE]
  > La Configuration de sécurité renforcée d’Internet Explorer (ESC) doit être désactivée pour que les compléments web Office fonctionnent. Si vous utilisez un ordinateur Windows Server comme votre client lors du développement des compléments, notez qu’ESC est activée par défaut dans Windows Server.

- L’un des éléments suivants en tant que navigateur par défaut : Internet Explorer 11 ou la dernière version de Microsoft Edge, Chrome, Firefox ou Safari (Mac OS).
- Un éditeur HTML et JavaScript tel que le Bloc-notes, [Visual Studio et les outils de développement Office ](https://www.visualstudio.com/features/office-tools-vs) ou un outil de développement web tiers.

## <a name="client-requirements-os-x-desktop"></a>Exigences en matière de client : ordinateur de bureau OS X

Outlook sur Mac, qui est distribué dans le cadre de Microsoft 365, prend en charge les compléments Outlook. l’exécution de compléments Outlook dans Outlook sur Mac a les mêmes conditions requises qu’Outlook sur Mac lui-même : le système d’exploitation doit être au moins le se X v 10.10 « Yosemite ». Comme Outlook sur Mac utilise WebKit comme moteur de disposition pour restituer les pages de complément, il n’existe pas de dépendance de navigateur supplémentaire.

Les versions de client minimales d’Office pour Mac prenant en charge les compléments Office sont les suivantes.

- Version Word 15.18 (160109)
- Version Excel 15.19 (160206)
- Version PowerPoint 15.24 (160614)

## <a name="client-requirements-browser-support-for-office-web-clients-and-sharepoint"></a>Exigences en matière de client : prise en charge du navigateur pour les clients web Office et SharePoint

Tout navigateur prenant en charge ECMAScript 5.1, HTML5 et CSS3, tel qu’Internet Explorer 11 ou la dernière version de Microsoft Edge, Chrome, Firefox ou Safari (Mac OS).


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>Exigences en matière de client : smartphones et tablettes autres que Windows

L’installation du logiciel suivant est nécessaire pour tester et exécuter les compléments Outlook. Ce logiciel est conçu spécialement pour les versions d’Outlook exécutées dans des navigateurs sur smartphones et tablettes utilisant d’autres systèmes d’exploitation que Windows.


| Application hôte | Appareil | Système d’exploitation | Compte Exchange | Navigateur mobile |
|:-----|:-----|:-----|:-----|:-----|
|Outlook sur Android|Tablettes et smartphones Android|Android KitKat 4.4 et version ultérieure|Sur la dernière mise à jour de Microsoft 365 Apps for Business ou Exchange Online|Application native pour Android, navigateur non applicable|
|Outlook sur iOS|Tablettes iPad, smartphones iPhone|iOS 11 ou version ultérieure|Sur la dernière mise à jour de Microsoft 365 Apps for Business ou Exchange Online|Application native pour iOS, navigateur non applicable|
|Outlook sur le web|iPhone 4, iPad 2, iPod Touch 4 (ou version ultérieure de ces appareils)|iOS 5 ou version ultérieure|Sur Microsoft 365, Exchange Online ou local sur Exchange Server 2013 ou version ultérieure|Safari|

> [!NOTE]
> Les applications natives OWA pour Android, OWA pour iPad et OWA pour iPhone ont été [supprimées](https://support.office.com/article/Microsoft-OWA-mobile-apps-are-being-retired-076ec122-4576-4900-bc26-937f84d25a4b) et ne sont plus requises ou disponibles pour les tests des compléments Outlook.


## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
- [Disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md)
- [Navigateurs utilisés par les compléments Office](browsers-used-by-office-web-add-ins.md)
