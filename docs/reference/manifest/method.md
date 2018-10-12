# <a name="method-element"></a>Élément Method

Spécifie une méthode individuelle de l’interface API JavaScript pour Office requise pour l’activation de votre complément Office.

**Type de complément :** contenu, volet Office

## <a name="syntax"></a>Syntaxe

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>Contenu dans

[Méthodes](methods.md)

## <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|Name|string|obligatoire|Spécifie le nom de la méthode qualifiée requise avec son objet parent. Par exemple, pour spécifier la méthode **getSelectedDataAsync**, vous devez spécifier `"Document.getSelectedDataAsync"`.|

## <a name="remarks"></a>Remarques

Les éléments **Methods** et **Method** ne sont pas pris en charge dans les compléments de courrier. Pour plus d’informations sur les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

> [!IMPORTANT] 
> Comme il n’existe aucun moyen pour spécifier la version minimale de condition requise pour les différentes méthodes, pour vous assurer qu’une méthode est disponible à l’exécution, vous devez également utiliser une instruction **if** lors de l’appel de cette méthode dans le script de votre complément. Pour plus d’informations, voir [Présentation de l’interface API JavaScript pour Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

