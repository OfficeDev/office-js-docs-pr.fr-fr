---
title: Fonctionnalité d’envoi des compléments Outlook
description: Permet de traiter un élément ou d’empêcher les utilisateurs d’effectuer certaines actions. Permet aussi aux compléments de définir certaines propriétés pendant l’envoi.
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 7dc35d71f78ec282ecedda9eba0ab4fc52811fee
ms.sourcegitcommit: 77617f6ad06e07f5ff8078b26301748f73e2ee01
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/29/2020
ms.locfileid: "44413188"
---
# <a name="on-send-feature-for-outlook-add-ins"></a>Fonctionnalité d’envoi des compléments Outlook

La fonctionnalité d’envoi des compléments Outlook vous permet de traiter un élément de message ou réunion, ou d’empêcher les utilisateurs d’effectuer certaines actions. Elle permet aussi aux compléments de définir certaines propriétés pendant l’envoi. Par exemple, vous pouvez utiliser la fonctionnalité d’envoi pour :

- Empêcher un utilisateur d’envoyer des informations sensibles ou de laisser la ligne d’objet vide.  
- Ajouter un destinataire spécifique à la ligne CC dans les messages ou à la ligne destinataires facultatifs des réunions.

La fonctionnalité d’envoi est déclenchée par le type d’événement `ItemSend` et est sans interface utilisateur.

Pour en savoir plus sur les limites de la fonctionnalité d’envoi, consultez la section [Limites](#limitations) plus loin dans cet article.

## <a name="supported-clients-and-platforms"></a>Clients et plateformes pris en charge

Le tableau suivant présente les combinaisons client-serveur prises en charge pour la fonctionnalité d’envoi. Les combinaisons exclues ne sont pas prises en charge.

| Client | Exchange Online | Exchange 2016 en local<br>(Mise à jour cumulative 6 ou version ultérieure) | Exchange 2019 en local<br>(Mise à jour cumulative 1 ou ultérieure) |
|---|:---:|:---:|:---:|
|Windows :<br>version 1910 (Build 12130,20272) ou version ultérieure|Oui|Oui|Oui|
|CMA<br>Build 16,30 ou version ultérieure|Oui|Non|Non|
|Navigateur Web :<br>interface utilisateur d’Outlook moderne|Oui|Non applicable|Non applicable|
|Navigateur Web :<br>interface utilisateur Outlook classique|Non applicable|Oui|Oui|

> [!NOTE]
> La fonctionnalité d’envoi a été publiée dans l’ensemble de conditions requises 1,8 (pour plus d’informations, consultez la rubrique [Current Server and client support](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) ).

> [!IMPORTANT]
> Les compléments qui utilisent la fonctionnalité d’envoi ne sont pas autorisés dans [AppSource](https://appsource.microsoft.com).

## <a name="how-does-the-on-send-feature-work"></a>Comment marche la fonctionnalité d’envoi ?

Vous pouvez utiliser la fonctionnalité d’envoi pour créer un complément Outlook qui intègre l’événement synchrone `ItemSend`. Cet événement détecte le moment où l’utilisateur clique sur le bouton **Envoyer**(ou le bouton **Envoyer mise à jour** pour les réunions existantes) et peut servir à bloquer l’envoi de l’élément s’il n’est pas validé. Par exemple, quand un utilisateur déclenche un événement d’envoi de message, un complément Outlook qui utilise la fonctionnalité d’envoi peut :

- Lire et valider le contenu du message
- Vérifier que la ligne d’objet du message est remplie
- Définir un destinataire prédéterminé

La validation est effectuée côté client dans Outlook lorsque l’événement Send est déclenché, et le complément dispose d’un maximum de 5 minutes avant que celui-ci n’expire. Si la validation échoue, l’envoi de l’élément est bloqué et un message d’erreur s’affiche dans une barre d’informations qui invite l’utilisateur à agir.

La capture d’écran suivante montre une barre d’informations invitant l’expéditeur à renseigner l’objet du message.

<br/>

![Capture d’écran qui montre un message d’erreur invitant l’utilisateur à renseigner l’objet du message](../images/block-on-send-subject-cc-inforbar.png)

<br/>

<br/>

La capture d’écran suivante montre une barre d’informations informant l’expéditeur que des mots bloqués ont été trouvés.

<br/>

![Capture d’écran qui montre un message d’erreur indiquant que des mots bloqués ont été trouvés](../images/block-on-send-body.png)

## <a name="limitations"></a>Limites

Les limites de la fonctionnalité d’envoi sont les suivantes.

- **AppSource** &ndash; Vous ne pouvez pas publier de compléments Outlook qui utilisent la fonctionnalité d’envoi sur [AppSource](https://appsource.microsoft.com). car ils ne seront pas validés par AppSource. Les compléments qui utilisent la fonctionnalité d’envoi doivent être déployés par les administrateurs.
- **Manifeste** &ndash; Le complément prend en charge un seul événement `ItemSend`. Si votre manifeste comprend plusieurs événements `ItemSend`, il ne sera pas validé.
- **Performances**&ndash; : plusieurs allers-retours vers le serveur web hébergeant le complément peuvent nuire aux performances du complément. Imaginez alors ce qu’occasionnerait la création de compléments nécessitant plusieurs opérations de messagerie ou réunions.
- **Envoyer plus tard** (Mac uniquement) &ndash; S’il y a des compléments d’envoi, la fonctionnalité **Envoyer plus tard** n’est pas disponible.

### <a name="mailbox-typemode-limitations"></a>Limites concernant le type ou le mode de boîte aux lettres

La fonctionnalité d’envoi est uniquement prise en charge pour les boîtes aux lettres utilisateur dans Outlook sur le web, sur Windows et sur Mac. La fonctionnalité n’est pas prise en charge pour les types et modes de boîte aux lettres suivants.

- Boîtes aux lettres partagées\*
- Boîtes aux lettres de groupe
- Mode hors connexion

Outlook bloque l’envoi si la fonctionnalité d’envoi est activée pour ces scénarios de boîtes aux lettres. Toutefois, si un utilisateur répond à un e-mail dans une boîte aux lettres de groupe, le complément d’envoi n’est pas exécuté et le message est envoyé.

> [!IMPORTANT]
> \*La fonctionnalité d’envoi doit fonctionner sur des boîtes aux lettres ou des dossiers partagés si le complément [implémente également la prise en charge des scénarios d’accès délégué](delegate-access.md).

## <a name="multiple-on-send-add-ins"></a>Compléments d’envoi multiples

Si plusieurs compléments d’envoi sont installés, ils s’exécutent dans l’ordre dans lequel ils sont reçus par les API `getAppManifestCall` ou `getExtensibilityContext`. Si le premier complément autorise l’envoi du message, le deuxième complément peut modifier un paramètre qui le bloque. Par contre, le premier complément n’est pas réexécuté si les autres compléments installés autorisent l’envoi.

Par exemple, Complément1 et Complément2 utilisent la fonctionnalité d’envoi. Complément1 est installé en premier, et Complément2 en deuxième. Complément1 vérifie que le mot Fabrikam apparaît dans le message pour autoriser l’envoi.  À l’inverse, Complément2 supprime toutes les occurrences du mot Fabrikam. Le message est alors envoyé sans le mot Fabrikam (à cause de l’ordre d’installation de Complément1 et Complément2).

## <a name="deploy-outlook-add-ins-that-use-on-send"></a>Déployer des compléments Outlook qui utilisent la fonctionnalité d’envoi

Nous recommandons aux administrateurs de déployer les compléments Outlook qui utilisent la fonctionnalité d’envoi. Les administrateurs doivent vérifier que le complément d’envoi :

- est présent lors de l’ouverture d’un élément de composition (pour les e-mails : nouveau message, répondre ou transférer).
- ne peut pas être fermé ou désactivé par l’utilisateur.

## <a name="install-outlook-add-ins-that-use-on-send"></a>Installer des compléments Outlook qui utilisent la fonctionnalité d’envoi

Dans Outlook, la fonctionnalité d’envoi exige la configuration des compléments en fonction des types d’événement d’envoi. Sélectionnez la plateforme que vous voulez configurer.

### <a name="web-browser---classic-outlook"></a>[Navigateur web – Outlook classique](#tab/classic)

Les compléments Outlook (classique) sur le web qui utilisent la fonctionnalité d’envoi s’exécutent pour les utilisateurs auxquels une stratégie de boîte aux lettres Outlook sur le web est attribuée, dont la valeur *OnSendAddinsEnabled* est définie sur **True**.

Pour installer un nouveau complément, exécutez les cmdlets Exchange Online PowerShell suivantes.

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> Pour découvrir comment utiliser PowerShell à distance afin de se connecter à Exchange Online, consultez la rubrique [Connexion à Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).

#### <a name="enable-the-on-send-feature"></a>Activer la fonctionnalité d’envoi

Par défaut, la fonctionnalité d’envoi est désactivée. Les administrateurs peuvent activer la fonctionnalité d’envoi en exécutant les cmdlets Exchange Online PowerShell.

Pour activer les compléments d’envoi pour tous les utilisateurs :

1. Créez une stratégie de boîte aux lettres Outlook sur le web.

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > Les administrateurs peuvent utiliser une stratégie existante, mais la fonctionnalité d’envoi est uniquement prise en charge sur certains types de boîtes aux lettres. La fonctionnalité d’envoi est bloquée par défaut sur les boîtes aux lettres non prises en charge dans Outlook sur le web.

2. Activez la fonctionnalité d’envoi.

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. Attribuez la stratégie à des utilisateurs.

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-feature-for-a-group-of-users"></a>Activer la fonctionnalité d’envoi pour un groupe d’utilisateurs

Pour activer la fonctionnalité d’envoi pour un groupe spécifique d’utilisateurs, suivez les étapes ci-dessous.  Dans cet exemple, un administrateur souhaite uniquement activer un complément d’envoi Outlook sur le web dans un environnement réservé aux utilisateurs du service financier.

1. Créez une stratégie de boîte aux lettres Outlook sur le web pour le groupe.

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > Les administrateurs peuvent utiliser une stratégie existante, mais la fonctionnalité d’envoi est uniquement prise en charge sur certains types de boîtes aux lettres (pour en savoir plus, consultez la section [Limites concernant le type de boîte aux lettres](#multiple-on-send-add-ins) plus haut dans cet article). La fonctionnalité d’envoi est bloquée par défaut sur les boîtes aux lettres non prises en charge dans Outlook sur le web.

2. Activez la fonctionnalité d’envoi.

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. Attribuez la stratégie à des utilisateurs.

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> vous devez attendre 60 minutes avant que la stratégie prenne effet. Sinon, redémarrez Internet Information Services (IIS). Une fois la stratégie prise en compte, la fonctionnalité d’envoi est activée pour le groupe.

#### <a name="disable-the-on-send-feature"></a>Désactiver la fonctionnalité d’envoi

Pour désactiver la fonctionnalité d’envoi pour un utilisateur ou affecter une stratégie de boîte aux lettres Outlook sur le web dont l’indicateur est désactivé, exécutez les cmdlets suivantes. Dans cet exemple, la stratégie de boîte aux lettres est *ContosoCorpOWAPolicy*.

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> Pour en savoir plus sur l’utilisation de la cmdlet **Set-OwaMailboxPolicy** en vue de configurer des stratégies de boîte aux lettres Outlook sur le web existantes, consultez la rubrique [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).

Pour désactiver la fonctionnalité d’envoi pour tous les utilisateurs auxquels une stratégie de boîte aux lettres Outlook sur le web spécifique est attribuée, exécutez les cmdlets suivantes.

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="web-browser---modern-outlook"></a>[Navigateur web – Outlook moderne](#tab/modern)

Les compléments pour Outlook sur le web (moderne) qui utilisent la fonctionnalité d’envoi doivent s’exécuter pour tous les utilisateurs qui les ont installés. Toutefois, si les utilisateurs sont obligés d’exécuter le complément pour respecter les normes de conformité, la stratégie de boîte aux lettres doit avoir l’indicateur *OnSendAddinsEnabled* défini sur **true**.

Pour installer un nouveau complément, exécutez les cmdlets Exchange Online PowerShell suivantes.

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> Pour découvrir comment utiliser PowerShell à distance afin de se connecter à Exchange Online, consultez la rubrique [Connexion à Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).

#### <a name="disable-the-on-send-policy"></a>Désactiver la stratégie d’envoi

Par défaut, la stratégie d’envoi est activée. Pour désactiver la stratégie d’envoi pour un utilisateur ou affecter une stratégie de boîte aux lettres Outlook sur le web dont l’indicateur est désactivé, exécutez les cmdlets suivantes. Dans cet exemple, la stratégie de boîte aux lettres est *ContosoCorpOWAPolicy*.

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> Pour en savoir plus sur l’utilisation de la cmdlet **Set-OwaMailboxPolicy** en vue de configurer des stratégies de boîte aux lettres Outlook sur le web existantes, consultez la rubrique [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).

Pour désactiver la stratégie d’envoi pour tous les utilisateurs auxquels une stratégie de boîte aux lettres Outlook sur le web spécifique est attribuée, exécutez les cmdlets suivantes.

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

#### <a name="enable-the-on-send-policy"></a>Activer la stratégie d’envoi

Les administrateurs peuvent activer la fonctionnalité d’envoi en exécutant les cmdlets Exchange Online PowerShell.

Pour activer les compléments d’envoi pour tous les utilisateurs :

1. Créez une stratégie de boîte aux lettres Outlook sur le web.

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > Les administrateurs peuvent utiliser une stratégie existante, mais la fonctionnalité d’envoi est uniquement prise en charge sur certains types de boîtes aux lettres. La fonctionnalité d’envoi est bloquée par défaut sur les boîtes aux lettres non prises en charge dans Outlook sur le web.

2. Activez la fonctionnalité d’envoi.

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. Attribuez la stratégie à des utilisateurs.

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-policy-for-a-group-of-users"></a>Activer la stratégie d’envoi pour un groupe d’utilisateurs

Pour activer la stratégie d’envoi pour un groupe spécifique d’utilisateurs, suivez les étapes ci-dessous.  Dans cet exemple, un administrateur souhaite uniquement activer une stratégie de complément d’envoi Outlook sur le web dans un environnement réservé aux utilisateurs du service financier.

1. Créez une stratégie de boîte aux lettres Outlook sur le web pour le groupe.

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > Les administrateurs peuvent utiliser une stratégie existante, mais la fonctionnalité d’envoi est uniquement prise en charge sur certains types de boîtes aux lettres (pour en savoir plus, consultez la section [Limites concernant le type de boîte aux lettres](#multiple-on-send-add-ins) plus haut dans cet article). La fonctionnalité d’envoi est bloquée par défaut sur les boîtes aux lettres non prises en charge dans Outlook sur le web.

2. Activez la stratégie d’envoi.

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. Attribuez la stratégie à des utilisateurs.

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> vous devez attendre 60 minutes avant que la stratégie prenne effet. Sinon, redémarrez Internet Information Services (IIS). Une fois la stratégie prise en compte, la fonctionnalité d’envoi est appliquée pour le groupe.

### <a name="windows"></a>[Windows](#tab/windows)

Les compléments pour Outlook sur Windows qui utilisent la fonctionnalité d’envoi doivent s’exécuter pour tous les utilisateurs qui les ont installés. Toutefois, si les utilisateurs sont obligés d’exécuter le complément pour respecter les normes de conformité, la stratégie de groupe **Désactiver l’envoi lorsque les extensions Web ne peuvent pas être chargées** doit être **Activée** sur chaque ordinateur concerné.

Pour activer les stratégies de boîte aux lettres, les administrateurs peuvent télécharger [l’outil modèles d’administration](https://www.microsoft.com/download/details.aspx?id=49030) et accéder aux modèles d’administration les plus récents en exécutant l’éditeur de stratégie de groupe local, **gpedit.msc**.

#### <a name="what-the-policy-does"></a>Rôle de la stratégie

Pour des raisons de conformité, il se peut que les administrateurs doivent s’assurer que les utilisateurs ne peuvent pas envoyer de d’éléments message ou réunion tant que la dernière mise à jour du complément n’est pas disponible. Les administrateurs doivent activer la stratégie de groupe **Désactiver l’envoi lorsque les extensions Web ne peuvent pas être chargées**, de sorte que tous les compléments sont mis à jour à partir d’Exchange et disponibles pour vérifier que chaque élément message ou réunion respecte les règles et réglementations attendues lors de l’envoi.

|État de la stratégie|Résultat|
|---|---|
|Désactivé|Envoi autorisé. L’élément message ou réunion peut être envoyé sans exécution du complément sur envoi, même si le complément n’a pas encore été mis à jour à partir d’Exchange.|
|Activé|L’envoi est autorisé uniquement lorsque le complément a été mis à jour à partir d’Exchange ; dans le cas contraire, l’envoi est bloqué.|

#### <a name="manage-the-on-send-policy"></a>Gérer la stratégie d’envoi

Par défaut, la stratégie d’envoi est désactivée. Les administrateurs peuvent activer la stratégie d’envoi en veillant à ce que le paramètre de la stratégie de groupe de l’utilisateur **Désactiver l'envoi lorsque les extensions Web ne sont pas chargées** soit **Activé**. Pour désactiver la stratégie pour un utilisateur, l’administrateur doit la paramétrer sur **Désactivé**. Pour gérer ce paramètre de stratégie, vous pouvez procéder comme suit.

1. Téléchargez l’[outil de modèles d’administration](https://www.microsoft.com/download/details.aspx?id=49030).
1. Ouvrez l’éditeur de stratégie de groupe local (**gpedit.msc**).
1. Accédez à **Configuration utilisateur > modèles d’administration > Microsoft Outlook 2016 > Sécurité > Centre de gestion de la confidentialité**.
1. Sélectionnez le paramètre **Désactiver l’envoi lorsque les extensions Web ne peuvent pas charger**.
1. Ouvrir le lien pour modifier le paramètre de stratégie.
1. Dans la fenêtre de dialogue **Désactiver l’envoi lorsque les extensions Web ne peuvent pas charger**, sélectionnez **Activée** ou **Désactivée**, puis sélectionnez **OK** ou **Appliquer** pour appliquer la mise à jour.

### <a name="mac"></a>[Mac](#tab/unix)

Les compléments pour Outlook sur Mac qui utilisent la fonctionnalité d’envoi doivent s’exécuter pour tous les utilisateurs qui les ont installés. Toutefois, si les utilisateurs sont obligés d’exécuter le complément pour respecter les normes de conformité, le paramètre de boîte aux lettres suivant doit être appliqué sur l’ordinateur de chaque utilisateur. Ce paramètre ou cette clé sont compatibles avec CFPreference, ce qui signifie qu’elle peut être définie à l’aide d’un logiciel de gestion d’entreprise pour Mac, tel que Jamf Pro.

|||
|:---|:---|
|**Domaine**|com.microsoft.outlook|
|**Clé**|OnSendAddinsWaitForLoad|
|**Type de données**|Valeur booléenne|
|**Valeurs possibles**|false (par défaut)<br>true|
|**Disponibilité**|16.27|
|**Commentaires**|Cette clé crée une stratégie onSendMailbox.|

#### <a name="what-the-setting-does"></a>Le rôle du paramètre

Pour des raisons de conformité, il se peut que les administrateurs doivent s’assurer que les utilisateurs ne peuvent pas envoyer de d’éléments message ou réunion tant que la dernière mise à jour des compléments n’est pas disponible. Les administrateurs doivent activer la clé **OnSendAddinsWaitForLoad**, de sorte que tous les compléments sont mis à jour à partir d’Exchange et disponibles pour vérifier que chaque élément message ou réunion respecte les règles et réglementations attendues lors de l’envoi.

|État de la clé|Résultat|
|---|---|
|false|Envoi autorisé. L’élément message ou réunion peut être envoyé sans exécution du complément sur envoi, même si le complément n’a pas encore été mis à jour à partir d’Exchange.|
|true|L’envoi est autorisé uniquement lorsque les compléments ont été mis à jour à partir d’Exchange ; dans le cas contraire, l’envoi est bloqué et le bouton **Envoyer** est désactivé.|

---

## <a name="on-send-feature-scenarios"></a>Scénarios de la fonctionnalité d’envoi

Voici tous les scénarios pris en charge et non pour les compléments qui utilisent la fonctionnalité d’envoi.

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-but-no-add-ins-are-installed"></a>La fonctionnalité d’envoi est activée sur la boîte aux lettres de l’utilisateur, mais aucun complément n’est installé.

Dans ce scénario, l’utilisateur peut envoyer des éléments message ou réunion sans l’exécution des compléments.

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-and-add-ins-that-supports-on-send-are-installed-and-enabled"></a>La fonctionnalité d’envoi est activée sur la boîte aux lettres de l’utilisateur et les compléments qui prennent en charge cette fonctionnalité sont installés et activés

Les compléments s’exécutent pendant l’événement d’envoi pour autoriser ou empêcher l’utilisateur d’envoyer son message.

### <a name="mailbox-delegation-where-mailbox-1-has-full-access-permissions-to-mailbox-2"></a>Délégation de boîte aux lettres, où la Boîte aux lettres 1 dispose des autorisations d’accès total à la Boîte aux lettres 2

#### <a name="web-browser-classic-outlook"></a>Navigateur web (Outlook classique)

|Scénario|Fonctionnalité d’envoi (Boîte aux lettres 1)|Fonctionnalité d’envoi (Boîte aux lettres 2)|Session web Outlook (classique)|Résultat|Pris en charge ?|
|:------------|:------------|:--------------------------|:---------|:-------------|:-------------|
|1 |Activé|Activé|Nouvelle session|La boîte aux lettres 1 ne peut pas envoyer un message ou un élément de réunion provenant de la boîte aux lettres 2.|N’est pas pris en charge actuellement. Pour y remédier, utilisez le scénario 3.|
|n°2|Désactivé|Activé|Nouvelle session|La boîte aux lettres 1 ne peut pas envoyer un message ou un élément de réunion provenant de la boîte aux lettres 2.|N’est pas pris en charge actuellement. Pour y remédier, utilisez le scénario 3.|
|3|Activé|Activé|Même session|Les compléments d’envoi attribués à la boîte aux lettres 1 exécutent la fonctionnalité d’envoi.|Pris en charge.|
|4 |Activé|Désactivé|Nouvelle session|Aucun complément d’envoi ne s’exécute ; un message ou un élément de réunion est envoyé.|Pris en charge.|

#### <a name="web-browser-modern-outlook-windows-mac"></a>Navigateur web (Outlook moderne), Windows, Mac

Pour appliquer l’envoi, les administrateurs doivent s’assurer que la stratégie a été activée sur les deux boîtes aux lettres. Pour plus d’informations sur la prise en charge de l’accès délégué dans un complément, voir [Activer les scénarios d’accès délégué dans un complément Outlook](delegate-access.md).

### <a name="group-1-is-a-modern-group-mailbox-and-user-mailbox-1-is-a-member-of-group-1"></a>Le Groupe 1 est une boîte aux lettres de groupe moderne et la Boîte aux lettres d’utilisateur 1 est membre du Groupe 1

<br/>

|Scénario|Stratégie d’envoi de la boîte aux lettres 1|Compléments d’envoi activés ?|Action de la boîte aux lettres 1|Résultat|Pris en charge ?|
|:------------|:-------------------------|:-------------------|:---------|:----------|:-------------|
|1 |Activé|Oui|La Boîte aux lettres 1 compose un nouveau message ou réunion pour le Groupe 1.|Les compléments d’envoi s’exécutent pendant l’envoi.|Oui|
|n°2|Activé|Oui|La boîte aux lettres 1 compose un nouveau message ou réunion pour le Groupe 1, dans la fenêtre du Groupe 1 dans Outlook sur le web.|Les compléments d’envoi ne s’exécutent pas pendant l’envoi.|N’est pas pris en charge actuellement. Pour y remédier, utilisez le scénario 1.|

### <a name="user-mailbox-with-on-send-add-in-featurepolicy-enabled-add-ins-that-support-on-send-are-installed-and-enabled-and-offline-mode-is-enabled"></a>La fonctionnalité/stratégie d’envoi est activée sur la boîte aux lettres de l’utilisateur, les compléments qui prennent en charge cette fonctionnalité sont installés et activés et le mode hors connexion est activé

Les compléments d’envoi s’exécutent en fonction de l’état en ligne de l’utilisateur, du serveur principal du complément et d’Exchange.

#### <a name="users-state"></a>État de l’utilisateur

Les compléments d’envoi s’exécutent pendant l’envoi, si l’utilisateur est en ligne. Si l’utilisateur est hors ligne, les compléments d’envoi ne s’exécutent pas pendant l’envoi et l’élément message ou réunion n’est pas envoyé.

#### <a name="add-in-backends-state"></a>État du serveur de complément

Un complément sur envoi s’exécute si son serveur principal est en ligne et joignable. Si le serveur principal est hors connexion, l’envoi est désactivé.

#### <a name="exchanges-state"></a>État d’Exchange

Les compléments d’envoi s’exécutent pendant l’envoi, si le serveur Exchange est en ligne et joignable. Si le complément sur envoi ne peut pas accéder à Exchange et que la stratégie ou l’applet de commande applicable sont activés, l’envoi est désactivé.

> [!NOTE]
> Sur Mac en mode hors connexion, le bouton **Envoyer** (ou le bouton **Envoyer mise à jour** pour les réunions existantes) est désactivé et une notification indique que l’organisation n’autorise pas l’envoi lorsque l’utilisateur est hors connexion.


## <a name="code-examples"></a>Exemples de code

Les exemples de code ci-dessous vous montrent comment créer un complément d’envoi simple. Pour télécharger l’exemple de code sur lequel se basent ces exemples, consultez l’article [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).

### <a name="manifest-version-override-and-event"></a>Manifeste, remplacement de version et événement

L’exemple de code [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) comprend deux manifestes :

- `Contoso Message Body Checker.xml` &ndash; : montre comment vérifier la présence de mots non autorisés ou d’informations sensibles dans le corps d’un message pendant l’envoi.  

- `Contoso Subject and CC Checker.xml` &ndash; : montre comment ajouter un destinataire à la ligne Cc et vérifier que le message comporte une ligne d’objet pendant l’envoi.  

Dans le fichier manifeste `Contoso Message Body Checker.xml`, insérez le fichier de fonction et le nom de la fonction qui doit être appelée lors d’un événement `ItemSend`. L’opération s’exécute de façon synchrone.

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this case, the function validateBody will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
                <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateBody" />
            </ExtensionPoint>
        </DesktopFormFactor>
    </Host>
</Hosts>
```

> [!IMPORTANT]
> Si vous utilisez Visual Studio 2019 pour développer votre complément d’envoi, il se peut que vous obteniez un avertissement de validation semblable à celui-ci : « il s’agit d’un xsi : type incorrect http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events . » Pour contourner ce message, vous avez besoin d’une version plus récente de la MailAppVersionOverridesV1_1. xsd qui a été fournie comme GitHub dans un [blog sur cet avertissement](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).

Pour le fichier manifeste `Contoso Subject and CC Checker.xml`, l’exemple suivant montre le fichier de fonction et le nom de la fonction à appeler dans l’événement d’envoi du message.

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this case the function validateSubjectAndCC will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
                <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateSubjectAndCC" />
            </ExtensionPoint>
        </DesktopFormFactor>
    </Host>
</Hosts>
```

<br/>

L’API d’envoi nécessite `VersionOverrides v1_1`. L’exemple vous montre comment ajouter le nœud `VersionOverrides` dans votre manifeste.

```xml
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <!-- On Send requires VersionOverridesV1_1 -->
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
         ...
     </VersionOverrides>
</VersionOverrides>
```

> [!NOTE]
> Pour plus d’informations, voir les commandes suivantes :
> - [Manifestes de complément Outlook](manifests.md)
> - [Manifeste XML des compléments Office](../overview/add-in-manifests.md)


### <a name="event-and-item-objects-and-bodygetasync-and-bodysetasync-methods"></a>Les objets `Event` et `item` et les méthodes `body.getAsync` et `body.setAsync`

Pour accéder au message ou élément de réunion sélectionné (dans cet exemple, le message que vous venez de composer), utilisez l’espace de noms `Office.context.mailbox.item`. L’événement `ItemSend` est automatiquement transmis via la fonctionnalité d’envoi vers la fonction spécifiée dans le manifeste &mdash;,dans cet exemple, la fonction `validateBody`.

```js
var mailboxItem;

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
}

// Entry point for Contoso Message Body Checker add-in before send is allowed.
// <param name="event">ItemSend event is automatically passed by on-send code to the function specified in the manifest.</param>
function validateBody(event) {
    mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);
}
```

Le corps actuel de la fonction `validateBody` s’affiche dans le format spécifié (HTML) et transmet l’objet « event » `ItemSend` auquel le code souhaite accéder avec la méthode du rappel. En plus de la méthode `getAsync`, l’objet `Body` fournit également une méthode `setAsync` utile pour remplacer le corps du message par le texte spécifié.

> [!NOTE]
> Pour en savoir plus, consultez les articles relatifs à l’objet [Event](/javascript/api/office/office.addincommands.event) et à la méthode [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).
  

### <a name="notificationmessages-object-and-eventcompleted-method"></a>Objet `NotificationMessages` et méthode `event.completed`

La fonction `checkBodyOnlyOnSendCallBack` utilise une expression régulière pour déterminer si le corps du message contient des mots bloqués. Si elle trouve une correspondance dans un tableau de mots bloqués, il bloque l’envoi du message et avertit l’expéditeur via la barre d’informations. Pour ce faire, il utilise la propriété `notificationMessages` de l'objet `Item` pour renvoyer un objet `NotificationMessages`. Il ajoute ensuite une notification à l’élément en appelant la méthode `addAsync`, comme illustré dans l’exemple suivant.

```js
// Determine whether the body contains a specific set of blocked words. If it contains the blocked words, block email from being sent. Otherwise allow sending.
// <param name="asyncResult">ItemSend event passed from the calling function.</param>
function checkBodyOnlyOnSendCallBack(asyncResult) {
    var listOfBlockedWords = new Array("blockedword", "blockedword1", "blockedword2");
    var wordExpression = listOfBlockedWords.join('|');

    // \b to perform a "whole words only" search using a regular expression in the form of \bword\b.
    // i to perform case-insensitive search.
    var regexCheck = new RegExp('\\b(' + wordExpression + ')\\b', 'i');
    var checkBody = regexCheck.test(asyncResult.value);

    if (checkBody) {
        mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked words have been found in the body of this email. Please remove them.' });
        // Block send.
        asyncResult.asyncContext.completed({ allowEvent: false });
    }

    // Allow send.
    asyncResult.asyncContext.completed({ allowEvent: true });
}
```

Voici les paramètres pour la méthode `addAsync` :

- `NoSend` &ndash; : chaîne correspondant à une clé spécifiée par un développeur pour référencer un message de notification. Vous pouvez l’utiliser pour modifier ce message ultérieurement. La clé ne peut pas contenir plus de 32 caractères.
- `type`&ndash; : l’une des propriétés du paramètre d’objet JSON. Représente le type d’un message ; les types correspondent aux valeurs de l’énumération [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype). Les valeurs possibles sont Indicateur de progression, Message d’information ou Message d’erreur. Dans cet exemple, `type` est un message d’erreur.  
- `message`&ndash; : l’une des propriétés du paramètre d’objet JSON. Dans cet exemple, `message` correspond au texte du message de notification.

Pour signaler que le complément a terminé le traitement de l’événement `ItemSend` déclenché par l’opération d’envoi, appelez la méthode `event.completed({allowEvent:Boolean})`. La propriété `allowEvent` est une valeur booléenne. Si la valeur est définie sur `true`, l’envoi est autorisé. Si la valeur est définie sur `false`, l’envoi du message est bloqué.

> [!NOTE]
> Pour plus d’informations, consultez les articles relatifs à [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) et à [completed](/javascript/api/office/office.addincommands.event).

### <a name="replaceasync-removeasync-and-getallasync-methods"></a>Méthodes `replaceAsync`, `removeAsync` et `getAllAsync`

En plus de la méthode `addAsync`, l'objet `NotificationMessages` inclut également les méthodes `replaceAsync`, `removeAsync` et `getAllAsync`.  Ces méthodes ne sont pas utilisées dans cet exemple de code.  Pour plus d’informations, consultez l’article relatif à l’objet [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).


### <a name="subject-and-cc-checker-code"></a>Code vérificateur de l’objet et de la ligne CC

L’exemple de code suivant vous montre comment ajouter un destinataire à la ligne Cc et vérifier que le message comporte un objet pendant l’envoi. Cet exemple utilise la fonctionnalité d’envoi pour autoriser ou interdire l’envoi d’un e-mail.  

```js
// Invoke by Contoso Subject and CC Checker add-in before send is allowed.
// <param name="event">ItemSend event is automatically passed by on-send code to the function specified in the manifest.</param>
function validateSubjectAndCC(event) {
    shouldChangeSubjectOnSend(event);
}

// Determine whether the subject should be changed. If it is already changed, allow send. Otherwise change it.
// <param name="event">ItemSend event passed from the calling function.</param>
function shouldChangeSubjectOnSend(event) {
    mailboxItem.subject.getAsync(
        { asyncContext: event },
        function (asyncResult) {
            addCCOnSend(asyncResult.asyncContext);
            //console.log(asyncResult.value);
            // Match string.
            var checkSubject = (new RegExp(/\[Checked\]/)).test(asyncResult.value)
            // Add [Checked]: to subject line.
            subject = '[Checked]: ' + asyncResult.value;

            // Determine whether a string is blank, null, or undefined.
            // If yes, block send and display information bar to notify sender to add a subject.
            if (asyncResult.value === null || (/^\s*$/).test(asyncResult.value)) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Please enter a subject for this email.' });
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // If can't find a [Checked]: string match in subject, call subjectOnSendChange function.
                if (!checkSubject) {
                    subjectOnSendChange(subject, asyncResult.asyncContext);
                    //console.log(checkSubject);
                }
                else {
                    // Allow send.
                    asyncResult.asyncContext.completed({ allowEvent: true });
                }
            }
        });
}

// Add a CC to the email. In this example, CC contoso@contoso.onmicrosoft.com
// <param name="event">ItemSend event passed from calling function</param>
function addCCOnSend(event) {
    mailboxItem.cc.setAsync(['Contoso@contoso.onmicrosoft.com'], { asyncContext: event });
}

// Determine whether the subject should be changed. If it is already changed, allow send, otherwise change it.
// <param name="subject">Subject to set.</param>
// <param name="event">ItemSend event passed from the calling function.</param>
function subjectOnSendChange(subject, event) {
    mailboxItem.subject.setAsync(
        subject,
        { asyncContext: event },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Unable to set the subject.' });

                // Block send.
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // Allow send.
                asyncResult.asyncContext.completed({ allowEvent: true });
            }
        });
}
```

Pour savoir comment ajouter un destinataire à la ligne Cc et vérifier que le message comporte une ligne d’objet pendant l’envoi, et découvrir les API disponibles, consultez l’article relatif à l’exemple [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send). Le code est accompagné de commentaires détaillés.

## <a name="see-also"></a>Voir aussi

- [Présentation de l’architecture et des fonctionnalités des compléments Outlook](outlook-add-ins-overview.md)
- [Démonstration de la commande du complément Outlook](https://github.com/OfficeDev/outlook-add-in-command-demo)
