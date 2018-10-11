# <a name="allowsnapshot-element"></a>Élément AllowSnapshot

Indique si une capture instantanée de votre complément de contenu est enregistrée avec le document hôte.

**Type de complément :** Contenu

## <a name="syntax"></a>Syntaxe

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Remarques

 > [!IMPORTANT]
 > **AllowSnapshot** est `true` par défaut. Cela crée une image du complément visible pour les utilisateurs qui ouvrent le document dans une version de l’application hôte ne prenant pas en charge les compléments Office, ou fournissant une image statique du complément si l’application hôte ne peut pas se connecter au serveur qui héberge le complément. Toutefois, cela signifie également que les informations potentiellement sensibles affichées dans le complément sont accessibles directement à partir du document hébergeant le complément.

