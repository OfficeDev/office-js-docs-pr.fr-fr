# <a name="sourcelocation-element"></a>Élément SourceLocation

Spécifie les emplacements des fichiers source pour votre extension Office sous forme d’URL comprenant entre 1 et 2 018 caractères. L’emplacement source doit être une adresse HTTPS, et non un chemin d’accès de fichier.

**Type d'extension :** contenu, volet Office, courrier

## <a name="syntax"></a>Syntaxe

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>Contenu dans

- [DefaultSettings](defaultsettings.md) (compléments de contenu et extensions du volet Office)
- [FormSettings](formsettings.md) (extensions pour courrier)
- [ExtensionPoint](extensionpoint.md) (extensions pour courriers contextuels)

## <a name="can-contain"></a>Peut contenir

[remplacement](override.md)

## <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**requis**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|requis|Spécifie la valeur par défaut de ce paramètre pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).|
