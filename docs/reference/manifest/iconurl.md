# <a name="iconurl-element"></a>Élément IconUrl

Spécifie l’URL de l’image utilisée pour représenter votre complément Office dans l’UX d’insertion et l’Office Store.

**Type de complément :** contenu, volet Office, courrier

## <a name="syntax"></a>Syntaxe

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>Peut contenir

[Override](override.md)

## <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|string|obligatoire|Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Remarques

Pour un complément de messagerie, l’icône s’affiche dans l’interface utilisateur, sous **Fichier**  >  **Gérer les compléments** (Outlook) ou sous **Paramètres**  >  **Gérer les compléments** (Outlook Web App). Pour un complément de contenu ou de volet Office, l’icône s’affiche dans l’interface utilisateur, sous **Insérer**  >  **Compléments**. Pour tous les types de compléments, l’icône est également utilisée sur le site de l’Office Store si vous publiez votre complément dans l’Office Store.

L’image doit être dans un des formats de fichier suivants : GIF, JPG, PNG, EXIF, BMP ou TIFF. Pour les applications de volet de tâches et de contenu, l’image spécifiée doit être de 32 x 32 pixels. Pour les applications de messagerie, l’image doit être de 64 x 64 pixels. Il est recommandé de spécifier également une icône pour les applications hôte Office exécutées sur des écrans à DPI élevé, à l’aide de l’élément [HighResolutionIconUrl](highresolutioniconurl.md). Pour plus d’informations, voir la section _Créer une identité visuelle cohérente pour votre application_ dans [Créer des référencements efficaces dans AppSource et dans Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).
