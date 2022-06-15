---
title: Configuration requise pour exécuter des compléments Office
description: Découvrez les exigences du client et du serveur qu’un utilisateur final doit exécuter Office compléments.
ms.date: 06/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 06699e8a2c498eb6ad2f9832a8369beef5af4786
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091033"
---
# <a name="requirements-for-running-office-add-ins"></a>Configuration requise pour exécuter des compléments Office

Cet article décrit la configuration logicielle et matérielle requise pour l’exécution des compléments Office.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

Pour obtenir une vue générale de l’emplacement de prise en charge des compléments Office, consultez [Office disponibilité des applications clientes et de la plateforme pour Office compléments](/javascript/api/requirement-sets).

## <a name="server-requirements"></a>Exigences en matière de serveur

Pour pouvoir installer et exécuter des Complément Office, vous devez d’abord déployer les fichiers manifeste et de pages web pour l’interface utilisateur et le code de votre complément sur les emplacements de serveur appropriés.

Pour tous les types de complément (compléments de contenu, Outlook et volet Office, et les commandes de compléments), vous devez déployer les fichiers de pages web de votre complément sur un serveur web ou un service d’hébergement web, tel que [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> lorsque vous développez et déboguez un complément dans Visual Studio, Visual Studio déploie et exécute les fichiers de page web de votre complément localement avec IIS Express et ne nécessite aucun serveur web supplémentaires.

Pour les compléments de contenu et de volet Office, dans les applications clientes Office prises en charge ( Excel, PowerPoint, Project ou Word ), vous avez également besoin d’un [catalogue d’applications](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) sur SharePoint pour charger le fichier manifeste XML du complément, ou vous devez déployer le complément à l’aide [d’applications intégrées](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps).

Pour tester et exécuter un complément Outlook, le compte de messagerie Outlook de l’utilisateur doit résider sur Exchange 2013 ou une version ultérieure, qui est disponible via Microsoft 365, Exchange Online ou via une installation locale. L’utilisateur ou l’administrateur installe les fichiers manifeste pour les compléments Outlook sur ce serveur.

> [!NOTE]
> Les comptes de messagerie POP et IMAP dans Outlook ne prennent pas en charge les compléments Office.

## <a name="client-requirements-windows-desktop-and-tablet"></a>Exigences en matière de client : ordinateur de bureau et tablette Windows

Le logiciel suivant est requis pour développer un complément Office pour les clients de bureau ou les clients web Office pris en charge qui s’exécutent sur Windows ordinateurs de bureau, ordinateurs portables ou tablettes.

- Pour les ordinateurs de bureau Windows x86 et x64 et les tablettes telles que Surface Pro :
  - La version 32 bits ou 64 bits d’Office 2013 ou une version ultérieure s’exécutant sur Windows 7 ou une version ultérieure.
  - Excel 2013, Outlook 2013, PowerPoint 2013, Project Professionnel 2013, Project 2013 SP1 ou Word 2013, ou une version ultérieure du client Office, si vous testez ou exécutez un Complément Office, notamment pour l’un de ces clients de bureau Office. Les clients de bureau Office peuvent être installés sur site ou par le biais de « Démarrer en un clic » sur l’ordinateur client.

  Si vous disposez d’un abonnement Microsoft 365 valide et que vous n’avez pas accès au client Office, vous pouvez [télécharger et installer la dernière version de Office](https://support.microsoft.com/office/4414eaaf-0478-48be-9c42-23adc4716658).

- Microsoft Edge doit être installé, mais ne doit pas nécessairement être le navigateur par défaut. Pour prendre en charge Office compléments, le client Office qui agit en tant qu’hôte utilise des composants de navigateur qui font partie de Microsoft Edge.

  > [!NOTE]
  >
  > - À proprement parler, il est possible de développer des compléments sur un ordinateur sur lequel Internet Explorer 11 est installé, mais pas Microsoft Edge. Toutefois, Internet Explorer est utilisé pour exécuter des compléments uniquement sur certaines combinaisons plus anciennes de Windows et de versions Office. Pour plus d’informations, voir [Navigateurs utilisés par les compléments Office](browsers-used-by-office-web-add-ins.md). Nous vous déconseillons d’utiliser des environnements tels que votre environnement de développement de complément principal. Toutefois, si vous êtes susceptible d’avoir des clients de votre complément qui fonctionnent dans ces combinaisons plus anciennes, nous vous recommandons de prendre en charge Internet Explorer. Pour plus d’informations, consultez [Support Internet Explorer 11](../develop/support-ie-11.md).
  > - La Configuration de sécurité renforcée d’Internet Explorer (ESC) doit être désactivée pour que les compléments web Office fonctionnent. Si vous utilisez un ordinateur Windows Server comme votre client lors du développement des compléments, notez qu’ESC est activée par défaut dans Windows Server.

- L’un des éléments suivants en tant que navigateur par défaut : Internet Explorer 11 ou la dernière version de Microsoft Edge, Chrome, Firefox ou Safari (Mac OS).
- Un éditeur HTML et JavaScript tel que [Visual Studio Code](https://code.visualstudio.com/), [Visual Studio et les outils](https://www.visualstudio.com/features/office-tools-vs) de développement Microsoft, ou un outil de développement web non Microsoft.

## <a name="client-requirements-os-x-desktop"></a>Exigences en matière de client : ordinateur de bureau OS X

Outlook sur Mac, qui est distribué dans le cadre de Microsoft 365, prend en charge Outlook compléments. L’exécution Outlook compléments dans Outlook sur Mac a les mêmes exigences que Outlook sur Mac lui-même : le système d’exploitation doit être au moins OS X v10.10 « Yosemite ». Comme Outlook sur Mac utilise WebKit comme moteur de disposition pour restituer les pages de complément, il n’existe pas de dépendance de navigateur supplémentaire.

Les versions de client minimales d’Office pour Mac prenant en charge les compléments Office sont les suivantes.

- Version Word 15.18 (160109)
- Version Excel 15.19 (160206)
- Version PowerPoint 15.24 (160614)

## <a name="client-requirements-browser-support-for-office-web-clients-and-sharepoint"></a>Exigences en matière de client : prise en charge du navigateur pour les clients web Office et SharePoint

Tout navigateur, à l’exception d’Internet Explorer, qui prend en charge ECMAScript 5.1, HTML5 et CSS3, comme Microsoft Edge, Chrome, Firefox ou Safari (Mac OS).

## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>Exigences du client : Smartphone et tablette non Windows

Spécifiquement pour Outlook s’exécutant sur des smartphones et des tablettes non Windows, les logiciels suivants sont nécessaires pour tester et exécuter Outlook compléments.

| Application Office | Appareil | Système d’exploitation | Compte Exchange | Navigateur mobile |
|:-----|:-----|:-----|:-----|:-----|
|Outlook sur Android|- Android tablettes<br>- smartphones Android|- Android 4.4 KitKat ou version ultérieure|Sur la dernière mise à jour de Applications Microsoft 365 pour les PME ou Exchange Online|Navigateur non applicable. Utilisez l’application native pour Android.<sup> 1</sup>|
|Outlook sur iOS|- iPad tablettes<br>- smartphones iPhone|- iOS 11 ou version ultérieure|Sur la dernière mise à jour de Applications Microsoft 365 pour les PME ou Exchange Online|Navigateur non applicable. Utilisez l’application native pour iOS.<sup> 1</sup>|
|Outlook sur le web (moderne)<sup>2</sup>|- iPad 2 ou version ultérieure<br>- Android tablettes |- iOS 5 ou version ultérieure<br>- Android 4.4 KitKat ou version ultérieure|Sur Microsoft 365, Exchange Online|- Microsoft Edge<br>- Chrome<br>- Firefox<br>- Safari|
|Outlook sur le web (classique)|- iPhone 4 ou version ultérieure<br>- iPad 2 ou version ultérieure<br>- iPod Touch 4 ou version ultérieure|- iOS 5 ou version ultérieure|Localement Exchange Server 2013 ou version ultérieure<sup>3</sup>|- Safari|

> [!NOTE]
> <sup>1</sup> OWA pour Android, OWA pour iPad et OWA pour iPhone applications natives ont été [déconseillés](https://support.microsoft.com/office/076ec122-4576-4900-bc26-937f84d25a4b).
>
> <sup>2</sup> Les Outlook sur le web modernes sur les smartphones iPhone et Android ne sont plus nécessaires ni disponibles pour tester Outlook compléments.
>
> <sup>3</sup> Les compléments ne sont pas pris en charge dans Outlook sur Android, sur iOS et le web mobile moderne avec des comptes Exchange locaux.

> [!TIP]
> Vous pouvez faire la distinction entre les deux versions d’Outlook, classique et moderne, dans un navigateur Web en regardant la barre d’outils de votre boîte aux lettres.
>
> **moderne**
>
> ![Capture d'écran partielle de la barre d'outils moderne d' Outlook.](../images/outlook-on-the-web-new-toolbar.png)
>
> **classique**
>
> ![Capture d’écran partielle de la barre d’outils Outlook classique.](../images/outlook-on-the-web-classic-toolbar.png)

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
- [Application cliente Office et disponibilité de la plateforme pour les compléments Office](/javascript/api/requirement-sets)
- [Navigateurs utilisés par les compléments Office](browsers-used-by-office-web-add-ins.md)
