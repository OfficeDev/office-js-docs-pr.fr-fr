# <a name="appdomains-element"></a>Élément AppDomains

Répertorie tout domaine supplémentaire, en plus du domaine spécifié dans l’élément SourceLocation, qui sera utilisé par votre complément Office pour charger des pages. Pour chaque domaine supplémentaire, indiquez un élément AppDomain.

 **Type de complément :** contenu, volet Office, courrier

## <a name="syntax"></a>Syntaxe

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Peut contenir

[AppDomain](appdomain.md)

## <a name="remarks"></a>Remarques

Par défaut, votre complément peut charger n’importe quelle page qui se trouve dans le même domaine que l’emplacement indiqué dans l’élément **SourceLocation**. Pour charger des pages qui ne sont pas dans le même domaine que le complément, spécifiez les domaines à l’aide des éléments **AppDomains** et **AppDomain**. Cet élément ne peut être laissé vide. 
