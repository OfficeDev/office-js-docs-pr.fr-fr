# <a name="outlook-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour Outlook

Les versions API requises pour les compléments Outlook sont indiquées à l’aide de l’élément [Requirements](/office/dev/add-ins/reference/manifest/requirements) dans leur [manifeste](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests). Les compléments Outlook contiennent toujours un élément [Set](/office/dev/add-ins/reference/manifest/set) avec un attribut `Name` défini sur `Mailbox` et un attribut `MinVersion` défini sur l’ensemble minimal de conditions requises de l’API qui prend en charge les scénarios du complément.

Par exemple, l’extrait de manifeste suivant indique l’ensemble minimal de conditions requises 1.1 :

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

Toutes les API Outlook appartiennent à l’`Mailbox`[ensemble de conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements). L’ensemble de conditions requises `Mailbox` possède plusieurs versions et chaque nouvel ensemble d’API publié appartient à une version supérieure de l’ensemble. L’ensemble d’API le plus récent n’est pas pris en charge par tous les clients Outlook, mais si un client Outlook prend en charge un ensemble de conditions requises, toutes les API comprises dans cet ensemble sont également prises en charge.

La définition d’une version minimale d’ensemble de conditions requises dans le manifeste permet de contrôler les clients Outlook dans lesquels le complément va apparaître. Si un client ne prend pas en charge l’ensemble minimal de conditions requises, il ne charge pas le complément. Par exemple, si la version de l’ensemble de conditions requises spécifiée est 1.3, le complément n’apparaîtra pas dans les clients Outlook qui ne prennent pas en charge au minimum la version 1.3.

## <a name="using-apis-from-later-requirement-sets"></a>Utilisation des API d’un ensemble de conditions requises ultérieure

La définition d’un ensemble de conditions requises ne limite pas votre complément à utiliser les API de cette version. Par exemple, si le complément indique l’ensemble de conditions requises 1.1, mais qu’il s’est exécuté dans un client Outlook prenant en charge la version 1.3, le complément peut utiliser les API de l’ensemble de conditions requises 1.3.

Pour utiliser des API plus récentes, les développeurs peuvent simplement vérifier leur disponibilité en utilisant la technique JavaScript standard :

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

Ces vérifications ne sont pas nécessaires pour les API présentes dans l’ensemble de conditions requises dont la version est la même que celle spécifiée dans le manifeste.

## <a name="choosing-a-minimum-requirement-set"></a>Choix d’un ensemble minimal de conditions requises

Les développeurs doivent utiliser l’ensemble de conditions requises le plus ancien qui contient l’ensemble d’API critique pour leur scénario, sans lequel le complément ne fonctionne pas.

## <a name="clients"></a>Clients

Les clients suivants prennent en charge des compléments Outlook.

| Client | Ensembles de conditions requises des API prises en charge |
| --- | --- |
| Outlook 2019 pour Windows | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2019 pour Mac | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2016 (Démarrer en un clic) pour Windows | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2016 (MSI) pour Windows | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook 2016 pour Mac | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook 2013 pour Windows | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook pour iPhone | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook pour Android | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook sur le web (Office 365 et Outlook.com) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook Web App (Exchange 2013 en local) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1) |
| Outlook Web App (Exchange 2016 en local) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3) |
| Outlook Web App (Exchange 2019 en local) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |

> [!NOTE]
> La prise en charge de la version 1.3 dans Outlook 2013 a été ajoutée dans le cadre de la [mise à jour du 8 décembre 2015 pour Outlook 2013 (KB3114349)](https://support.microsoft.com/kb/3114349). La prise en charge de la version 1.4 dans Outlook 2013 a été ajoutée dans le cadre de la [mise à jour du 13 septembre 2016 pour Outlook 2013 (KB3118280)](https://support.microsoft.com/help/3118280).
