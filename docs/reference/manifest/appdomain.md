# <a name="appdomain-element"></a>AppDomain, élément

Indique un domaine supplémentaire permettant de charger des pages dans la fenêtre du complément.

**Type de complément :** application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> La valeur de l’élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain<AppDomain>`).

## <a name="contained-in"></a>Contenu dans

[AppDomains](appdomains.md)

## <a name="remarks"></a>Remarques

Les éléments **AppDomain** sont utilisés pour indiquer les domaines supplémentaires autres que celui spécifié dans l’[élément SourceLocation](sourcelocation.md). Pour plus d’informations, reportez-vous au [manifeste XML de compléments Office](/office/dev/add-ins/develop/add-in-manifests).
