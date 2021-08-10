---
title: Élément Ressources dans le fichier manifest
description: L’élément Resources contient des icônes, des chaînes, des URL pour le nœud VersionOverrides.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 9fbf28790b717389bf7248babcb69f28d4226b77c3139415fb14bbd712690074
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57094143"
---
# <a name="resources-element"></a>Élément Resources

Contient des icônes, des chaînes et des URL pour le nœud [VersionOverrides](versionoverrides.md). Un élément de manifeste indique une ressource à l’aide de l’**Id** de la ressource. Cela permet de conserver une taille de manifeste raisonnable, surtout lorsque les ressources sont disponibles en plusieurs versions selon les paramètres régionaux. Un **Id** doit être unique au sein du manifeste et doit comporter 32 caractères au maximum.

Chaque ressource peut avoir plusieurs éléments enfants **Override** afin que vous puissiez définir une ressource différente pour un paramètre régional spécifique.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Type  |  Description  |
|:-----|:-----|:-----|
|  [Images](#images)            |  image   |  Fournit l’URL HTTPS de l’image d’une icône. |
|  **URL**                |  url     |  Fournit l’URL HTTPS. Une URL peut comporter jusqu’à 2 048 caractères. |
|  **ShortStrings** |  string  |  Texte pour les éléments **Label** et **Title**. Chaque élément **String** comporte 125 caractères au maximum.|
|  **LongStrings**  |  string  | Texte pour les attributs **Description**. Chaque **chaîne** comporte 250 caractères au maximum.|

> [!NOTE]
> Vous devez utiliser le protocole SSL (Secure Sockets Layer) pour toutes les URL dans les éléments **Image** et **Url**.

### <a name="images"></a>Des images

Chaque icône doit avoir trois **éléments Images,** un pour chacune des trois tailles obligatoires :

- 16x16
- 32x32
- 80x80

Les tailles supplémentaires suivantes sont également pris en charge, mais ne sont pas obligatoires.

- 20x20
- 24x24
- 40x40
- 48x48
- 64x64

> [!IMPORTANT]
>
> - Si cette image est l’icône représentative de votre application, voir Créer des listes efficaces dans [AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) et dans Office pour la taille et d’autres exigences.
> - Outlook doit pouvoir mettre en cache les ressources d’image pour des raisons de performances. Par conséquent, le serveur qui héberge une ressource d’image ne doit pas ajouter les directives CACHE-CONTROL à l’en-tête de réponse. Outlook remplacera alors automatiquement une image générique ou par défaut.

## <a name="resources-examples"></a>Exemples de ressources

```XML
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp32-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp80-icon_default.png" />
        </bt:Image>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
        </bt:Url>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residLabel" DefaultValue="GetData">
          <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residToolTip" DefaultValue="Get data for your document.">
          <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
        </bt:String>
      </bt:LongStrings>
    </Resources>
```

```xml
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER//blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/blue-80.png"/>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
    <!-- other URLs -->
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="groupLabel" DefaultValue="Add-in Demo">
      <bt:Override Locale="ar-sa" Value="<Localized text>" />
    </bt:String>
    <!-- Other short strings -->
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment.">
      <bt:Override Locale="ar-sa" Value="<Localized text>." />
    </bt:String>
    <!-- Other long strings -->
  </bt:LongStrings>
</Resources>
```
