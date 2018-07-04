unit gSysUtils;

interface

uses
  Variants
  ,uTob
  {$if not Defined(EAGLCLIENT)}
    ,DB
  {$ifend}
  ,HEnt1
  ,cbpErrorMsg
  ,Classes
  ;

type
  ArrayOfStrings = array of String;
  ArrayOfWideStrings = array of WideString;
  ArrayOfIntegers = array of Integer;
  ArrayOfDouble = array of Double;
  ArrayOfVariants = array of Variant;

	MyNewLine = (wNormal, wException, wExceptionLig, wInsertion, wDuplication); //TS¤TODO -> deprecated

  ///<summary>
  ///Class d'exception de base aux méthodes communes (surcharge de la cbpException)
  ///</summary>
  ///<value>
  ///Permet de gérer le message d'erreur remonté lors de l'encapsulation d'une 
  ///transaction.
  ///Permet également de positionner le V_Pgi.ioError
  ///</value>
  ExceptionG = class(cbpException)
  private
    procedure SaveStatus(const Msg: WideString; NotifyIOError: Boolean = False;
      const IOError: TIOErr = oeUnknown);
  public
    constructor Create(const cbpIDError, SourceLocation: WideString;
      Args: array of Const; NotifyIOError: Boolean = False;
      const IOError: TIOErr = oeUnknown); overload;
  end;

  ExceptionGSys = class(ExceptionG);
  EGSysObjectUnassigned = class(ExceptionGSys);
  EGSysDateFormatUnknown = class(ExceptionGSys);
  EGSysRoundingNbdecimalOffBounds = class(ExceptionGSys);
  EGSysRoundingPrecisionOffBounds = class(ExceptionGSys);
  EGSysRoundingWeightOffBounds = class(ExceptionGSys);
  EGSysRoundingMethodUnknown = class(ExceptionGSys);
  EGSysRoundingContextUnknown = class(ExceptionGSys);
  EGSysFeatureTagUnknown = class(ExceptionGSys);

  ///<summary>
  ///Interface exposant les méthodes nécessaire à l'extraction de données typées
  ///en fonction d'un identifiant de contexte
  ///</summary>
  ITypedGetter = interface
  ['{679CCD32-D01A-4F18-A3BC-F827805D4C6A}']
    function GetValue(const Id: String): Variant;
    function GetString(const Id: String): String;
    function GetInteger(const Id: String): Integer;
    function GetDouble(const Id: String): Double;
    function GetBoolean(const Id: String): Boolean;
    function GetDateTime(const Id: String): TDateTime;
    function GetObject(const Id: String): TObject;
  end;

  ///<summary>
  ///Fournit une bibliothèque d'utilitaires de gestion de chaînes de caractères
  ///</summary>
  TGStrUtils = class
  public
    ///<summary>
    ///Retourne la position de "SubStr" dans "Str"
    ///</summary>
    ///<param name="Str">Chaîne à analyser</param>
    ///<param name="SubStr">Chaîne à rechercher</param>
    ///<param name="ExactSubStr">Mot entier ou début de SubStr (Mot entier par défaut)</param>
    ///<param name="Separator">séparateur des tokens de "Str"</param>
    ///<returns>Retourne la position de "SubStr" dans "Str" (-1 si non trouvé)</returns>
    class function TokenIndexOf(Str: String; const SubStr: String;
      const ExactSubStr: Boolean = True; const Separator: String = ';'): Integer;

    ///<summary>
    ///Lecture du jeton n° "Index"
    ///</summary>
    ///<param name="Str">Chaîne à analyser</param>
    ///<param name="Index">indice du jeton à trouver</param>
    ///<param name="Separator">séparateur des tokens de "Str"</param>
    ///<returns>Retourne le jeton présent dans "Str" à l'indice demandé</returns>
    class function TokenReadFromIndex(Str: String; const Index: Integer;
      const Separator: String = ';'): String;

    ///<summary>
    ///Compte le nombre de jetons
    ///</summary>
    ///<param name="Str">Chaîne à analyser</param>
    ///<param name="Separator">séparateur des tokens de "Str"</param>
    ///<returns>Retourne le nombre de jetons</returns>
    class function TokenCount(Str: String; const Separator: String = ';'): Integer;

    ///<summary>
    ///Retourne le plus grand jeton de la chaîne
    ///</summary>
    ///<param name="Str">Chaîne à analyser</param>
    ///<param name="Separator">Séparateur de jetons</param>
    ///<returns>la plus longue donnée de "Str"</returns>
    class function TokenGetGreater(Str: String; const Separator: String = ';'): String;

    ///<summary>
    ///retourne les "count" caractères de "Str" en partant du début
    ///</summary>
    ///<param name="Str">Chaîne copier</param>
    ///<param name="Count">nombre de caractères à copier</param>
    ///<returns>la partie gauche de "Str" de longueur "count"</returns>
    class function CopyLeft(const Str: String; const Count: Integer): String; overload;
    class function CopyLeft(const Str: WideString; const Count: Integer): WideString; overload;

    ///<summary>
    ///Retourne les "count" caractères de "Str" en partant de la fin
    ///</summary>
    ///<param name="Str">Chaîne copier</param>
    ///<param name="Count">nombre de caractères à copier</param>
    ///<returns>la partie droite de "Str" de longueur "count"</returns>
    class function CopyRight(const Str: String; const Count: Integer): String; overload;
    class function CopyRight(const Str: WideString; const Count: Integer): WideString; overload;

    ///<summary>
    ///supprime "count" caractères de "Str" en partant de la fin
    ///</summary>
    ///<param name="Str">Chaîne à couper</param>
    ///<param name="Count">nombre de caractères à supprimer</param>
    class procedure DeleteRight(var Str: String; const Count: Integer);

    ///<summary>
    ///supprime "count" caractères de "Str" en partant du début
    ///</summary>
    ///<param name="Str">Chaîne à couper</param>
    ///<param name="Count">nombre de caractères à supprimer</param>
    class procedure DeleteLeft(var Str: String; const Count: Integer);

    ///<summary>
    ///Alimente un HTStrings (unicode) à l'aide des jetons présents dans "Str"
    ///</summary>
    ///<param name="Str">chaîne tokenisée</param>
    ///<param name="Result">HTStrings déjà assignée à alimenter</param>
    ///<param name="Separator">séparateur des tokens de "Str"</param>
    ///<exception cref="EGSysObjectUnassigned">Result vaut nil</exception>
    ///<remarks>
    ///Ne vide pas la HTStrings
    ///</remarks>
    class procedure StrToHStringList(Str: String; const Result: HTStrings; const Separator: String = ';');

    ///<summary>
    ///convertit une chaîne "tokenisée" en TStrings
    ///</summary>
    ///<param name="Str">chaîne tokenisée</param>
    ///<param name="Result">TStrings déjà assignée à alimenter</param>
    ///<param name="Separator">séparateur des tokens de "Str"</param>
    ///<exception cref="EGSysObjectUnassigned">Result vaut nil</exception>
    ///<seealso cref="TGStrUtils.StringListToStr"/seealso>
    ///<remarks>
    ///Ne vide pas la TStrings
    ///</remarks>
    class procedure StrToStringList(Str: String; const Result: TStrings; const Separator: String = ';');

    ///<summary>
    ///convertit une TStrings en chaîne tokenisée
    ///</summary>
    ///<param name="StringList">TStrings déjà assignée à alimenter</param>
    ///<param name="Separator">séparateur des tokens de "Str"</param>
    ///<exception cref="EGSysObjectUnassigned">StringList vaut nil</exception>
    ///<seealso cref="TGStrUtils.StrToStringList"/>
    class function StringListToStr(const StringList: TStrings; const Separator: String = ';'): String;

    ///<summary>
    ///convertit une chaîne de caractères tokenisée en un tableau de chaînes
    ///</summary>
    ///<param name="Str">chaîne tokenisée</param>
    ///<param name="Separator">séparateur des tokens de "ttr"</param>
    ///<returns>le tableau contenant chaque jeton de donnée de "Str"</returns>
    ///<seealso cref="TGStrUtils.ArrayToStr"/seealso>
    class function AsArray(const Str: String; const Separator: string): ArrayOfStrings; overload;

    ///<summary>
    ///convertit une chaîne de caractères tokenisée en un tableau de chaînes
    ///</summary>
    ///<param name="Str">chaîne tokenisée</param>
    ///<param name="Strings">Tableau à compléter</param>
    ///<param name="Separator">séparateur des tokens de "ttr"</param>
    ///<returns>le tableau contenant chaque jeton de donnée de "Str"</returns>
    ///<remarks>
    ///Attention, le tableau fourni est complété. Il n'est pas vidé !
    ///</remarks>
    ///<seealso cref="TGStrUtils.ArrayToStr"/seealso>
    class procedure AsArray(Str: String; var Strings: ArrayOfStrings; const Separator: String); overload;

    ///<summary>
    ///convertit un tableau typé en une chaîne tokenisée
    ///</summary>
    ///<param name="xxxxxxx">Tableau typé</param>
    ///<param name="Separator">séparateur pour la chaîne de retour</param>
    ///<returns>une chaîne tokenisée</returns>
    class function ArrayToStr(const Strings: ArrayOfStrings; const Separator: String): String; overload;
    class function ArrayToStr(const Values: ArrayOfVariants; const Separator: String): String; overload;
    class function ArrayToStr(const Integers: ArrayOfIntegers; const Separator: String): String; overload;

    ///<summary>
    ///Ajoute un élément dans un tableau dynamique
    ///</summary>
    ///<param name="Strings">Tableau à alimenter</param>
    ///<param name="StrItem">Chaîne à ajouter</param>
    ///<returns>l'index du tableau auquel la chaîne a été rajoutée</returns>
    class function ArrayAddString(var Strings: ArrayOfStrings; const StrItem: String): Integer;

    ///<summary>
    ///Renvoie l'index de l'élément recherché dans le tableau dynamique
    ///</summary>
    ///<param name="Strings">Tableau à analyser</param>
    ///<param name="StrItem">Chaîne recherchée</param>
    ///<returns>l'index du tableau correspondant à la chaîne recherchée. -1 si non trouvé</returns>
    class function ArrayIndexOfString(const Strings: ArrayOfStrings; const StrItem: String): Integer;

    ///<summary>
    ///extrait de la chaîne la partie contenue entre "[" & "]"
    ///</summary>
    ///<param name="Str">Chaîne à analyser</param>
    ///<returns>la portion entre [ & ]</returns>
    class function ExtractParameters(const Str: String): String;

    ///<summary>
    ///copie un Item d'une TStringList vers une autre
    ///</summary>
    ///<param name="Item">Item à copier</param>
    ///<param name="FromList">Liste source</param>
    ///<param name="ToList">Liste de destination</param>
    ///<param name="ByItemName">Recherche dans la source via le nom de l'item (si la liste est composée d'éléments : Name=Value)</param>
    ///<param name="ExcludeDoubloon">Ne copie pas de la source vers la destination si l'élément est déjà présent dans la destination</param>
    ///<param name="DeleteItemSource">Supprime l'élément de la source</param>
    ///<param name="ToIndex">Index de l'élément dans la destination</param>
    ///<returns>
    ///* 0 : item copié
    ///* <0 : item source inexistant;
    ///* >0 : item de destination déjà existant ; correspond à son index
    ///</returns>
    class function StringListCopyItem(const Item: String; FromList, ToList: HTStrings;
      ByItemName: Boolean; ExcludeDoubloon: Boolean = True; DeleteItemSource: Boolean = False;
      const ToIndex: Integer = -1): Integer;

    ///<summary>
    ///Permet d'extraire une section d'une chaîne de caractères façon .ini :
    ///[SECTION1] bla bla [SECTION2] bla2 bla2
    ///</summary>
    ///<param name="SectionName">Nom de la section à extraire sans les "[]"</param>
    ///<param name="Buffer">Chaîne source</param>
    ///<param name="ModifyBuffer">Autorise la méthode à modifier le Buffer fourni</param>
    ///<returns>le contenu de la section</returns>
    class function ExtractSection(const SectionName: String; var Buffer: String;
      const ModifyBuffer: Boolean): String;

    ///<summary>
    ///complète la chaîne de taille fixe à droite par la caractère fourni
    ///</summary>
    ///<param name="Str">chaîne à compléter</param>
    ///<param name="Count">taille fixe de la chaîne retournée</param>
    ///<param name="PadChar">caractère de bourrage</param>
    ///<returns>la chaîne de longueur fixe complétée par le caractère fourni</returns>
    ///<value>
    ///chaîne d'entrée : 'a'
    ///PadRight(s, 5, '-')
    ///chaîne retournée : 'a----'
    ///</value>
    class function PadRight(const Str: String; const Count: Integer; const PadChar: Char = ' '): String;

    ///<summary>
    ///complète la chaîne de taille fixe à gauche par la caractère fourni
    ///</summary>
    ///<param name="Str">chaîne à compléter</param>
    ///<param name="Count">taille fixe de la chaîne retournée</param>
    ///<param name="PadChar">caractère de bourrage</param>
    ///<returns>la chaîne de longueur fixe complétée par le caractère fourni</returns>
    ///<value>
    ///chaîne d'entrée : 'a'
    ///PadLeft(s, 5, '-')
    ///chaîne retournée : '----a'
    ///</value>
    class function PadLeft(const Str: String; const Count: Integer; const PadChar: Char = ' '): String;

    ///<summary>
    ///retire de la chaîne les " de début et fin
    ///</summary>
    ///<param name="Str">Chaîne à analyser</param>
    ///<returns>la chaîne sans les "</returns>
    class function ExtractQuotes(const Str: String): String;

    ///<summary>
    ///répète "Str" "Count" fois
    ///</summary>
    ///<param name="Str">Chaîne à répéter</param>
    ///<returns>la chaîne concaténant "count" fois "Str"</returns>
    ///<remarks>
    ///pour répliquer un caractère simple, utiliser la routine de SysUtils :
    ///StringOfChar(Caractère, Count)
    ///</remarks>
    class function StringRepeat(const Str: string; Count: Integer): String;

    ///<summary>
    ///convertit une chaîne au format RTF en chaîne de caractères brut
    ///</summary>
    ///<param name="RTF">Chaîne au format RTF</param>
    ///<param name="ReplaceLineFeedWithSpace">remplacer les retours à la ligne par des espaces</param>
    ///<param name="DoTrimLeft">supprime les espace en début</param>
    ///<param name="TrailAfter">si la chaîne dépasse la taille, la suite est remplacée par des "..."</param>
    ///<returns>la chaîne sans le formattage RTF</returns>
    class function RTFToStr(const RTF: String; ReplaceLineFeedWithSpace: Boolean = False;
      DoTrimLeft: Boolean = True; const TrailAfter: Integer = 0): String;

    ///<summary>
    ///Convertit une chaîne quelconque en une chaîne composée uniquement de caractères alpha numériques
    ///</summary>
    ///<param name="Str">Chaîne à analyser</param>
    ///<param name="CharsToIgnore">caractères supplémentaires à ignorer</param>
    ///<returns>la chaîne alphanumérique</returns>
    class function StrToAlphaNumStr(const Str: String; const CharsToIgnore: String = ''): String;

    ///<summary>
    ///compresse une chaîne (algo ZIP)
    ///</summary>
    ///<param name="Str">chaîne à compresser</param>
    ///<returns>la chaîne "zippée"</returns>
    class function Zip(const Str: String): String;

    ///<summary>
    ///décompresse une chaîne (algo ZIP)
    ///</summary>
    ///<param name="Str">chaîne compressée</param>
    ///<returns>la chaîne "dézippée"</returns>
    class function Unzip(const Str: String): String;
  end;

  { ********************************************************************************************** }

  TDayReturned = (drSameDay, drFirstDay, drLastDay);
  TFullDateFormat = (
                      fdfYear,              //= AAAA
                      fdfSemester,          //= AAAA-S
                      fdfTrimester,         //= AAAA-S-T
                      fdfMonth,             //= AAAA-S-T-MM
                      fdfDay,               //= AAAA-S-T-MM-JJ
                      fdfHour,              //= AAAA-S-T-MM-JJ-HH
                      fdfMinute,            //= AAAA-S-T-MM-JJ-HHMM
                      fdfSecond             //= AAAA-S-T-MM-JJ-HHMMSS
                    );
  TTime = type TDateTime;  { pour ne pas adhérer à Controls.pas }

  ///<summary>
  ///Fournit une bibliothèque d'utilitaires de gestion de dates
  ///</summary>
  TGDateUtils = class
  public
    ///<summary>
    ///Transforme une date + une heure en une chaine de caractère de format précisé dans l'appel
    ///</summary>
    ///<param name="ADate">Date de référence</param>
    ///<param name="AFormat">Format/précision demandé</param>
    ///<param name="Truncate">permet de tronquer le résultat à la précision demandée,
    ///et complète le jusqu'à la seconde avec la valeur d'initialisation de chaque élément</param>
    ///<returns>renvoie une chaîne formatée en fonction des paramètres demandés</returns>
    ///<value>
    ///Exemple en date du 14/07/2002 à 11 h 55 minutes et 23 secondes
    ///
    ///Format :
    ///
    ///AAAA                           -> 2002
    ///AAAA-S                         -> 2002-2
    ///AAAA-S-T                       -> 2002-2-3
    ///AAAA-S-T-MM                    -> 2002-2-3-07
    ///AAAA-S-T-MM-JJ                 -> 2002-2-3-07-14
    ///AAAA-S-T-MM-JJ-HH              -> 2002-2-3-07-14-11
    ///AAAA-S-T-MM-JJ-HHMM            -> 2002-2-3-07-14-1155
    ///AAAA-S-T-MM-JJ-HHMMSS          -> 2002-2-3-07-14-115523
    ///</value>
    class function FormatFullDate(const ADate: TDateTime;
      const AFormat: TFullDateFormat = fdfDay;
      const Truncate: Boolean = True): String;

    ///<summary>
    ///fonction permettant de convertir une chaîne en format pour l'utilisation de "FormatFullDate"
    ///</summary>
    ///<param name="StrFormat">Format sous la forme d'une chaîne</param>
    ///<returns>l'équivalent de type TFullDateFormat</returns>
    ///<exception cref="EGSysDateFormatUnknown">StrFormat non géré dans la méthode</exception>
    ///<seealso cref="TGDateUtils.FormatFullDate"/seealso>
    ///<seealso cref="TGDateUtils.FullDateFormatToStr"/seealso>
    class function StrToFullDateFormat(const StrFormat: String): TFullDateFormat;

    ///<summary>
    ///fonction permettant de convertir un format utilisé dans "FormatFullDate" en chaîne
    ///</summary>
    ///<param name="AFormat">Format de type TFullDateFormat</param>
    ///<returns>l'équivalent sous la forme d'une chaîne de caractères</returns>
    ///<exception cref="EGSysDateFormatUnknown">AFormat non implémenté dans la méthode</exception>
    ///<seealso cref="TGDateUtils.FormatFullDate"/seealso>
    ///<seealso cref="TGDateUtils.StrToFullDateFormat"/seealso>
    class function FullDateFormatToStr(const AFormat: TFullDateFormat): String;

    ///<summary>
    ///Additionne un nombre d'années à la date demandée
    ///</summary>
    ///<param name="ADate">Date de référence</param>
    ///<param name="Count ">nombre d'années</param>
    ///<param name="DayReturned">jour retourné dans l'année</param>
    ///<returns>Renvoie la date en initialisant le jour en fonction du paramètre demandé</returns>
    ///<value>
    ///* le 1er jour de l'année calculée     : DayReturned = drFirstDay
    ///* le dernier jour de l'année calculée : DayReturned = drLastDay
    ///* le même jour de l'année calculée    : DayReturned = drSameDay
    ///</value>
    class function AddYear(const ADate: TDateTime; const Count: Integer;
      const DayReturned: TDayReturned = drSameDay): TDateTime;

    ///<summary>
    ///Additionne un nombre de semestres à la date demandée
    ///</summary>
    ///<param name="ADate">Date de référence</param>
    ///<param name="Count ">nombre de semestres</param>
    ///<param name="DayReturned">jour retourné dans le semestre</param>
    ///<returns>Renvoie la date en initialisant le jour en fonction du paramètre demandé</returns>
    ///<value>
    ///* le 1er jour du semestre calculé     : DayReturned = drFirstDay
    ///* le dernier jour du semestre calculé : DayReturned = drLastDay
    ///* le même jour du semestre calculé    : DayReturned = drSameDay
    ///</value>
    class function AddSemester(const ADate: TDateTime; const Count: Integer;
      const DayReturned: TDayReturned = drSameDay): TDateTime;

    ///<summary>
    ///Additionne un nombre de trimestres à la date demandée
    ///</summary>
    ///<param name="ADate">Date de référence</param>
    ///<param name="Count ">nombre de trimestres</param>
    ///<param name="DayReturned">jour retourné dans le trimestres</param>
    ///<returns>Renvoie la date en initialisant le jour en fonction du paramètre demandé</returns>
    ///<value>
    ///* le 1er jour du trimestre calculé     : DayReturned = drFirstDay
    ///* le dernier jour du trimestre calculé : DayReturned = drLastDay
    ///* le même jour du trimestre calculé    : DayReturned = drSameDay
    ///</value>
    class function AddTrimester(const ADate: TDateTime; const Count: Integer;
      const DayReturned: TDayReturned = drSameDay): TDateTime;

    ///<summary>
    ///Additionne un nombre de mois à la date demandée
    ///</summary>
    ///<param name="ADate">Date de référence</param>
    ///<param name="Count ">nombre de mois</param>
    ///<param name="DayReturned">jour retourné dans le mois</param>
    ///<returns>Renvoie la date en initialisant le jour en fonction du paramètre demandé</returns>
    ///<value>
    ///* le 1er jour du mois calculé     : DayReturned = drFirstDay
    ///* le dernier jour du mois calculé : DayReturned = drLastDay
    ///* le même jour du mois calculé    : DayReturned = drSameDay
    ///</value>
    class function AddMonth(const ADate: TDateTime; const Count: Integer;
      const DayReturned: TDayReturned = drSameDay): TDateTime;

    ///<summary>
    ///Ajoute ou retire le nombre de semaines
    ///</summary>
    ///<param name="ADate">Date de référence</param>
    ///<param name="Count ">nombre de semaines</param>
    ///<param name="DayReturned">jour retourné dans la semaine</param>
    ///<returns>Renvoie la date en initialisant le jour en fonction du paramètre demandé</returns>
    ///<value>
    ///* se positionne en début de la nouvelle semaine : DayReturned = drFirstDay
    ///* se positionne en fin de la nouvelle semaine   : DayReturned = drLastDay
    ///</value>
    class function AddWeek(const ADate: TDateTime; const Count: Integer;
      const DayReturned: TDayReturned = drSameDay): TDateTime;

    ///<summary>
    ///Additionne un nombre de jour à la date demandée
    ///</summary>
    ///<param name="ADate">Date de référence</param>
    ///<param name="Count ">nombre de jours</param>
    ///<param name="HourReturned">heure retournée dans la journée</param>
    ///<returns>Renvoie la date en initialisant le jour en fonction du paramètre demandé</returns>
    ///<value>
    ///* la 1ère heure du jour calculé     : HourReturned = drFirstDay
    ///* la dernière heure du jour calculé : HourReturned = drLastDay
    ///* à l'heure du jour calculé         : HourReturned = drSameDay
    ///</value>
    class function AddDay(const ADate: TDateTime; const Count: Integer;
      const HourReturned: TDayReturned = drSameDay): TDateTime;

    ///<summary>
    ///GetYear / GetMonth & GetDay permettent d'extraire une partie de la date
    ///</summary>
    ///<param name="ADate">Date de référence</param>
    ///<returns>Renvoie la partie demandée</returns>
    class function GetYear(const ADate: TDateTime): Word;
    class function GetMonth(const ADate: TDateTime): Word;
    class function GetDay(const ADate: TDateTime): Word;

    ///<summary>
    ///GetHour / GetMinute / GetSecond & GetMilliSecond permettent d'extraire une partie de l'heure
    ///</summary>
    ///<param name="ATime">Heure de référence</param>
    ///<returns>Renvoie la partie demandée</returns>
    class function GetHour(const ATime: TTime): Word;
    class function GetMinute(const ATime: TTime): Word;
    class function GetSecond(const ATime: TTime): Word;
    class function GetMilliSecond(const ATime: TTime): Word;

    ///<summary>
    ///calcul le quantième d'une date
    ///</summary>
    ///<param name="ADate">Date de référence</param>
    ///<returns>le quantième de la date demandée</returns>
    class function GetDayNumber(const ADate: TDateTime): Integer;

    ///<summary>
    ///renvoie le nom complet du mois en fonction de son numéro
    ///</summary>
    ///<param name="Month">numéro du mois</param>
    ///<param name="UpCase1rst">première lettre en majuscule</param>
    ///<returns>le mois en toutes lettres</returns>
    class function MonthToStr(const Month: Word; const UpCase1rst: Boolean = False): HString;

    ///<summary>
    ///renvoie le numéro de la dernière semaine de l'année de la date fournie
    ///</summary>
    ///<param name="ADate">date de référence</param>
    ///<returns>un numéro de semaine</returns>
    class function GetLastWeek(const ADate: TDateTime): Word;

    ///<summary>
    ///renvoie le numéro de trimestre correspondant au mois fourni
    ///</summary>
    ///<param name="Month">numéro de mois</param>
    ///<returns>le numéro de trimestre</returns>
    class function MonthToTrimester(const Month: Word): Word;

    ///<summary>
    ///renvoie le numéro de trimestre correspondant à la date fournie
    ///</summary>
    ///<param name="ADate">date de référence</param>
    ///<returns>le numéro de trimestre</returns>
    class function GetTrimester(const ADate: TDateTime): Word;

    ///<summary>
    ///renvoie le numéro du premier mois du trimestre fourni
    ///</summary>
    ///<param name="Trimester">numéro de trimestre</param>
    ///<returns>le numéro de mois</returns>
    class function GetFirstMonthOfTrimester(const Trimester: Word): Word;

    ///<summary>
    ///renvoie le numéro de semestre correspondant à la date fournie
    ///</summary>
    ///<param name="ADate">date de référence</param>
    ///<returns>le numéro de semestre</returns>
    class function MonthToSemester(const Month: Word): Word;

    ///<summary>
    ///renvoie le numéro du premier mois du semestre fourni
    ///</summary>
    ///<param name="Semester">numéro de semestre</param>
    ///<returns>le numéro de mois</returns>
    class function GetFirstMonthOfSemester(const Semester: Word): Word;

    ///<summary>
    ///renvoie le numéro de semestre correspondant à la date fournie
    ///</summary>
    ///<param name="ADate">date de référence</param>
    ///<returns>le numéro de semestre</returns>
    class function GetSemester(const ADate: TDateTime): Word;

    ///<summary>
    ///permet de contrôler la validité d'une date au format chaîne de caractères
    ///</summary>
    ///<param name="ADate">date de référence sous la forme d'une chaîne de caractères</param>
    ///<returns>vrai en cas de non conformité de la date</returns>
    class function isInvalidDate(const ADate: String): Boolean;

    ///<summary>
    ///convertit un temps hh,ss en hh,cc (Heure Seconde en Heure centième)
    ///</summary>
    ///<param name="ATime">Temps au format double (1 heure 1/2 = 1.30)</param>
    ///<returns>le temps converti</returns>
    ///<seealso cref="TGDateUtils.HHCCToTime"/seealso>
    class function TimeToHHCC(const ATime: Double): Double;

    ///<summary>
    ///convertit un temps hh,cc en hh,ss (Heure centième en Heure seconde)
    ///</summary>
    ///<param name="ATime">Temps au format double (1 heure 1/2 = 1.50)</param>
    ///<returns>le temps converti</returns>
    ///<value>
    ///* Format = HH : ne renvoie que les heures
    ///* Format = MM : ne renvoie que les heures et les minutes
    ///* Format = SS : ne renvoie que les heures, les minutes et les secondes
    ///* Format = MS : renvoie les heures, les minutes, les secondes et les millisecondes
    ///</value>
    ///<seealso cref="TGDateUtils.HHCCToTime"/seealso>
    class function HHCCToTime(const ATime: Double; const Format: String = 'SS'): Double;
  end;

  { ********************************************************************************************** }

  TRoundingContext = (rcTaux, rcPrix, rcMont);
	TRoundingPrecision = (rp100000, rp10000, rp1000, rp100, rp10, rp1, rp01, rp001, rp0001, rp00001);
  TRoundingMethod = (rndLower, rndHigher, rndNearest);

  //deprecated
  TArrondiQuoi = (tTPMTaux, tTPMPrix, tTPMMont);
	TArrondiPrec = (tAPCentaineDeMillier, tAPDizaineDeMillier, tAPMillier, tAPCentaine, tAPDizaine, tAPUnite, tAPDizieme, tAPCentieme, tAPMillieme, tAPDixMillieme);
  TArrondiMeth = (tAMInferieure, tAMSuperieure, tAMPlusProche);
  //deprecated

  ///<summary>
  ///Fournit une bibliothèque d'utilitaires de gestion d'arrondi
  ///</summary>
  TGRounding = class
  private
    class procedure TraceWarning(const Msg: String);
  public
    ///<summary>
    ///Arrondir une quantité, un prix ou tout autre chose en fonction d'une précision d'arrondi et d'une méthode
    ///</summary>
    ///<param name="Value">Valeur à arrondir</param>
    ///<param name="RoundingPrecision">Précision demandée</param>
    ///<param name="RoundingMethod">Méthode d'arrondi</param>
    ///<returns>La valeur arrondie</returns>
    class function Round(const Value: Double; const RoundingPrecision: TRoundingPrecision;
      const RoundingMethod: TRoundingMethod): Double;

    ///<summary>
    ///Transforme une précision d'arrondi en un nombre de décimal d'arrondi
    ///</summary>
    ///<param name="RoundingPrecision">précision de l'arrondi (type ordinal)</param>
    ///<returns>nombre de décimales d'arrondi</returns>
    ///<value>
    ///* rp100000 retourne -5 (au centaine de millier)
    ///* rp1      retourne 0  (à l'unité)
    ///* rp00001  retourne 4  (au dix millième)
    ///</value>
    ///<exception cref="EGSysRoundingPrecisionOffBounds">
    ///RoundingPrecision n'est pas implémenté dans la méthode
    ///</exception>
    ///<seealso cref="TGRounding.NbDecimalToRoundingPrecision"/seealso>
    class function RoundingPrecisionToNbDecimal(const RoundingPrecision: TRoundingPrecision): Integer;

    ///<summary>
    ///Transforme une précision d'arrondi en un poids d'arrondi
    ///</summary>
    ///<param name="RoundingPrecision">précision de l'arrondi (type ordinal)</param>
    ///<returns>un poids d'arrondi</returns>
    ///<value>
    ///* rp100000 retourne 100000 (au centaine de millier)
    ///* rp1      retourne 1      (à l'unité)
    ///* rp00001  retourne 0.0001 (au dix millième)
    ///</value>
    ///<exception cref="EGSysRoundingPrecisionOffBounds">
    ///RoundingPrecision n'est pas implémenté dans la méthode
    ///</exception>
    ///<seealso cref="TGRounding.WeightToRoundingPrecision"/seealso>
    class function RoundingPrecisionToWeight(const RoundingPrecision: TRoundingPrecision): Double;

    ///<summary>
    ///Transforme le nombre de décimales d'arrondi sauvegardé dans la base
    ///en un type d'arrondi géré dans les écran et la méthode d'arrondi
    ///</summary>
    ///<param name="NbDecimal">nombre de décimales</param>
    ///<returns>précision de l'arrondi (type ordinal)</returns>
    ///<value>
    ///* -5 retourne rp100000 (au centaine de millier)
    ///*  0 retourne rp1      (à l'unité)
    ///*  4 retourne rp00001  (au dix millième)
    ///</value>
    ///<exception cref="EGSysRoundingNbdecimalOffBounds">
    ///NbDecimal n'est pas traductible en TRoundingPrecision car non géré dans la méthode
    ///</exception>
    ///<seealso cref="TGRounding.RoundingPrecisionToNbDecimal"/seealso>
    class function NbDecimalToRoundingPrecision(const NbDecimal: Integer): TRoundingPrecision;

    ///<summary>
    ///Transforme un poids d'arrondi sauvegardé dans la base
    ///en un type d'arrondi géré dans les écran et la méthode d'arrondi
    ///</summary>
    ///<param name="Weight">poids de l'arrondi</param>
    ///<returns>précision de l'arrondi (type ordinal)</returns>
    ///<value>
    ///* 100000 retourne rp100000 (au centaine de millier)
    ///* 1      retourne rp1      (à l'unité)
    ///* 0.0001 retourne rp00001  (au dix millième)
    ///</value>
    ///<exception cref="EGSysRoundingWeightOffBounds">
    ///Weight n'est pas traductible en TRoundingPrecision car non géré dans la méthode
    ///</exception>
    ///<seealso cref="TGRounding.RoundingPrecisionToWeight"/seealso>
    class function WeightToRoundingPrecision(const Weight: Double): TRoundingPrecision;

    ///<summary>
    ///Transforme une méthode d'arrondi sauvegardée dans la base
    ///en un type d'arrondi géré dans les écran et la méthode d'arrondi
    ///</summary>
    ///<param name="Str">méthode d'arrondi au format littéral</param>
    ///<returns>une méthode d'arrondi (type ordinal)</returns>
    ///<value>
    ///* 'I' = rmLower (arrondi à l'inférieur)
    ///* 'P' = rmNearest (arrondi au plus proche)
    ///* 'S' = rmHigher (arrondi au supérieur)
    ///</value>
    ///<exception cref="EGSysRoundingMethodUnknown">
    ///Str n'est pas traductible en TRoundingMethod car non géré dans la méthode
    ///</exception>
    ///<seealso cref="TGRounding.RoundingMethodToStr"/seealso>
    class function StrToRoundingMethod(const Str: String): TRoundingMethod;

    ///<summary>
    ///Transforme une méthode gérée dans les écran et la méthode d'arrondi
    ///en un type d'arrondi sauvegardé dans la base
    ///</summary>
    ///<param name="RoundingMethod">méthode d'arrondi au format ordinal</param>
    ///<returns>une méthode d'arrondi (type string)</returns>
    ///<value>
    ///* rmLower = 'I' (arrondi à l'inférieur)
    ///* rmNearest = 'P' (arrondi au plus proche)
    ///* rmHigher = 'S' (arrondi au supérieur)
    ///</value>
    ///<exception cref="EGSysRoundingMethodUnknown">
    ///RoundingMethod n'est pas traductible en string car non géré dans la méthode
    ///</exception>
    ///<seealso cref="TGRounding.StrToRoundingMethod"/seealso>
    class function RoundingMethodToStr(const RoundingMethod: TRoundingMethod): String;

    ///<summary>
    ///Transforme un contexte d'arrondi littéral
    ///en un contexte d'arrondi géré dans les écran et la méthode d'arrondi
    ///</summary>
    ///<param name="Str">contexte d'arrondi au format littéral</param>
    ///<returns>un contexte d'arrondi (type ordinal)</returns>
    ///<value>
    ///* 'T' = rcPMTaux (Taux)
    ///* 'P' = rcPMPrix (Prix)
    ///* 'M' = rcPMMont (Montant)
    ///</value>
    ///<exception cref="EGSysRoundingContextUnknown">
    ///Str n'est pas traductible en TRoundingContext car non géré dans la méthode
    ///</exception>
    ///<seealso cref="TGRounding.RoundingContextToStr"/seealso>
    class function StrToRoundingContext(const Str: String): TRoundingContext;

    ///<summary>
    ///Transforme un contexte d'arrondi géré dans les écran et la méthode d'arrondi
    ///en un contexte d'arrondi littéral
    ///</summary>
    ///<param name="RoundingContext">contexte d'arrondi au format ordinal</param>
    ///<returns>un contexte d'arrondi (type string)</returns>
    ///<value>
    ///* rcPMTaux = 'T' (Taux)
    ///* rcPMPrix = 'P' (Prix)
    ///* rcPMMont = 'M' (Montant)
    ///</value>
    ///<exception cref="EGSysRoundingContextUnknown">
    ///RoundingContext n'est pas traductible en TRoundingContext car non géré dans la méthode
    ///</exception>
    ///<seealso cref="TGRounding.StrToRoundingContext"/seealso>
    class function RoundingContextToStr(const RoundingContext: TRoundingContext): String;

    ///<summary>
    ///arrondi une valeur selon le multiple fourni
    ///</summary>
    ///<param name="Value">valeur à arrondir</param>
    ///<param name="MultipleOf">multiple</param>
    ///<returns>l'arrondi supérieur du élevé</returns>
    ///<value>
    ///utile par exemple pour le champ GA_QPCBCONSO
    ///ex. : 10 demandé pour un multilpe de 3 renvoi 12
    ///      10 / 3 = 3.3333 arrondi au supérieur = 4 > 4 * 3 = 12
    ///</value>
    class function RoundByMultipleOf(const Value, MultipleOf: Double): Double;

    class function ArrondiPrec2RoundingPrecision(const ArrondiPrec: TArrondiPrec): TRoundingPrecision;
    class function ArrondiMeth2RoundingMethod(const ArrondiMeth: TArrondiMeth): TRoundingMethod;
    class function ArrondiQuoi2RoundingContext(const ArrondiQuoi: TArrondiQuoi): TRoundingContext;
  end;

  ///<summary>
  ///Fournit une bibliothèque d'utilitaires divers sur les types simples
  ///</summary>
  TGComUtils = class
  public
    ///<summary>
    ///permet de savoir si le variant est de type entier
    ///</summary>
    ///<param name="Value">valeur de type variant à identifier</param>
    ///<returns>l'évaluation de ce variant</returns>
    class function isVarInt(const Value: Variant): Boolean;

    ///<summary>
    ///convertit un entier en chaîne de 2 caractères
    ///</summary>
    ///<param name="Value">entier à convertir</param>
    ///<returns>une chaîne de 2 caractères</returns>
    ///<value>
    ///ex :
    /// 2  > '02'
    /// 10 > '10'
    ///</value>
    class function IntToC2(const Value: Integer): String;

    ///<summary>
    ///retourne un tableau dynamique typé depuis un tableau ouvert
    ///</summary>
    ///<param name="ArrayData">Tableau ouvert</param>
    ///<returns>un ArrayOf"Type"</returns>
    ///<value>
    ///les tableau ouverts sont des tableaux passés en paramètres à des
    ///fonctions. inconvénient : SetLength() est inutilisable avec de type
    ///de tableau. Dans delphi, saisir "Array", sélectionner le mot puis F1
    ///lire la section "Paramètres tableau ouvert"
    ///</value>
    class function GetArrayOfStrings(const ArrayData: array of String): ArrayOfStrings;
    class function GetArrayOfVariants(const ArrayData: array of Variant): ArrayOfVariants;

    ///<summary>
    ///permet de convertir une chaîne de caractères en TActionFiche
    ///</summary>
    ///<param name="Action">type d'action en toutes lettres</param>
    ///<returns>
    ///* CONSULTATION = taConsult
    ///* CREATION     = taCreat
    ///* MODIFICATION = taModif (défaut)
    ///</returns>
    ///<seealso cref="TGCtrlsUtils.ActionToStr"/seealso>
    class function StrToAction(const Action: String): TActionFiche;

    ///<summary>
    ///permet de convertir un TActionFiche en chaîne de caractères
    ///</summary>
    ///<param name="Action">type d'action</param>
    ///<returns>
    ///* CONSULTATION = taConsult (défaut)
    ///* CREATION     = taCreat
    ///* MODIFICATION = taModif
    ///</returns>
    ///<seealso cref="TGCtrlsUtils.StrToAction"/seealso>
    class function ActionToStr(const Action: TActionFiche): String;

    ///<summary>
    ///en fonction d'un préfixe métier, permet de savoir si on a accès
    ///à la fonctionnalité via les fiches de gestion en modification
    ///</summary>
    ///<param name="Prefix">préfixe métier</param>
    class function JAiLeDroitGestion(const Prefix: String): Boolean;

    ///<summary>
    ///en fonction d'un préfixe métier, permet de savoir si on a accès
    ///à la fonctionnalité via les fiches de gestion en consultation
    ///</summary>
    ///<param name="Prefix">préfixe métier</param>
    class function JAiLeDroitConsult(const Prefix: String): Boolean;
  end;

  ///<summary>
  ///Fournit une bibliothèque d'utilitaires de gestion d'arguments
  ///Un argument étant une chaîne composée comme suit :
  ///'ID1=Value1;ID2=Value2;IDn=ValueN' avec ';' paramétrable
  ///</summary>
  TGArgUtils = class
  public
    ///<summary>
    ///identifie la position d'un argument dans une chaîne d'arguments
    ///</summary>
    ///<param name="Arguments">chaîne d'arguments</param>
    ///<param name="ArgumentID">argument recherché</param>
    ///<param name="Separator">séparateur d'arguments</param>
    ///<returns>l'indice de position de l'argument recherché</returns>
    ///<value>
    ///ex :
    /// IndexOf('ID1=Value1;ID2=Value2;IDn=ValueN', 'ID2') = 2
    /// IndexOf('ID1=Value1;ID2=Value2;IDn=ValueN', 'ID4') = -1
    ///</value>
    class function IndexOfID(const Arguments: String; const ArgumentID: String;
      const Separator: String = ';'): Integer;

    ///<summary>
    ///permet de savoir si un argument est présent dans une chaîne d'arguments
    ///</summary>
    ///<param name="Arguments">chaîne d'arguments</param>
    ///<param name="ArgumentID">argument recherché</param>
    ///<param name="Separator">séparateur d'arguments</param>
    class function Exists(const Arguments: String; const ArgumentID: String;
      const Separator: String = ';'): Boolean;

    ///<summary>
    ///Dans l'argument MY_FIELD=MY_VALUE, retourne MY_FIELD
    ///</summary>
    ///<param name="Token">Token à analyser</param>
    ///<returns>L'élément identifiant de la valeur d'un argument</returns>
    class function GetID(const Argument: String): String;

    ///<summary>
    ///renvoie la valeur d'un argument demandé au sein d'une chaîne d'arguments
    ///</summary>
    ///<param name="Arguments">chaîne d'arguments</param>
    ///<param name="ArgumentID">argument recherché</param>
    ///<param name="Separator">séparateur d'arguments</param>
    ///<param name="IgnoreCase">recherche ArgumentID en tenant compte de la casse</param>
    ///<returns>la valeur sous la forme d'un variant</returns>
    ///<seealso cref="GetString();GetInteger();GetDouble();GetDateTime();GetBoolean();GetTob()"/seealso>
    class function GetValue(const Arguments, ArgumentID: String;
      const Separator: String = ';'; IgnoreCase: Boolean = True): String;

    ///<summary>
    ///renvoie la valeur d'un argument demandé au sein d'une chaîne d'arguments
    ///</summary>
    ///<param name="Arguments">chaîne d'arguments</param>
    ///<param name="ArgumentID">argument recherché</param>
    ///<param name="Separator">séparateur d'arguments</param>
    ///<param name="IgnoreCase">recherche ArgumentID en tenant compte de la casse</param>
    ///<returns>la valeur sous la forme d'une chaîne de caractères (chaîne vide par défaut)</returns>
    ///<seealso cref="GetValue();GetInteger();GetDouble();GetDateTime();GetBoolean();GetTob()"/seealso>
    class function GetString(const Arguments, ArgumentID : String;
      const Separator: String = ';'; IgnoreCase: Boolean = True): String;

    ///<summary>
    ///renvoie la valeur d'un argument demandé au sein d'une chaîne d'arguments
    ///</summary>
    ///<param name="Arguments">chaîne d'arguments</param>
    ///<param name="ArgumentID">argument recherché</param>
    ///<param name="Separator">séparateur d'arguments</param>
    ///<param name="IgnoreCase">recherche ArgumentID en tenant compte de la casse</param>
    ///<returns>la valeur sous la forme d'un entier (0 par défaut)</returns>
    ///<seealso cref="GetValue();GetString();GetDouble();GetDateTime();GetBoolean();GetTob()"/seealso>
    class function GetInteger(const Arguments, ArgumentID : String;
      const Separator: String = ';'; IgnoreCase: Boolean = True): Integer;

    ///<summary>
    ///renvoie la valeur d'un argument demandé au sein d'une chaîne d'arguments
    ///</summary>
    ///<param name="Arguments">chaîne d'arguments</param>
    ///<param name="ArgumentID">argument recherché</param>
    ///<param name="Separator">séparateur d'arguments</param>
    ///<param name="IgnoreCase">recherche ArgumentID en tenant compte de la casse</param>
    ///<returns>la valeur sous la forme d'un numérique flottant (0 par défaut)</returns>
    ///<seealso cref="GetValue();GetString();GetInteger();GetDateTime();GetBoolean();GetTob()"/seealso>
    class function GetDouble(const Arguments, ArgumentID : String;
      const Separator: String = ';'; IgnoreCase: Boolean = True): Double;

    ///<summary>
    ///renvoie la valeur d'un argument demandé au sein d'une chaîne d'arguments
    ///</summary>
    ///<param name="Arguments">chaîne d'arguments</param>
    ///<param name="ArgumentID">argument recherché</param>
    ///<param name="Separator">séparateur d'arguments</param>
    ///<param name="IgnoreCase">recherche ArgumentID en tenant compte de la casse</param>
    ///<returns>la valeur sous la forme d'une date (01/01/1900 par défaut)</returns>
    ///<seealso cref="GetValue();GetString();GetDouble();GetDouble();GetBoolean();GetTob()"/seealso>
    class function GetDateTime(const Arguments, ArgumentID : String;
      const Separator: String = ';'; IgnoreCase: Boolean = True): tDateTime;

    ///<summary>
    ///renvoie la valeur d'un argument demandé au sein d'une chaîne d'arguments
    ///</summary>
    ///<param name="Arguments">chaîne d'arguments</param>
    ///<param name="ArgumentID">argument recherché</param>
    ///<param name="Separator">séparateur d'arguments</param>
    ///<param name="IgnoreCase">recherche ArgumentID en tenant compte de la casse</param>
    ///<returns>la valeur sous la forme d'un booléen : X=True, -=False (False par défaut)</returns>
    ///<seealso cref="GetValue();GetString();GetDouble();GetDouble();GetDateTime();GetTob()"/seealso>
    class function GetBoolean(const Arguments, ArgumentID : String;
      const Separator: String = ';'; IgnoreCase: Boolean = True): Boolean;

    ///<summary>
    ///renvoie la valeur d'un argument demandé au sein d'une chaîne d'arguments
    ///</summary>
    ///<param name="Arguments">chaîne d'arguments</param>
    ///<param name="ArgumentID">argument recherché</param>
    ///<param name="Separator">séparateur d'arguments</param>
    ///<param name="IgnoreCase">recherche ArgumentID en tenant compte de la casse</param>
    ///<returns>la valeur sous la forme d'une Tob (nil par défaut)</returns>
    ///<seealso cref="GetValue();GetString();GetDouble();GetDouble();GetBoolean();GetTob()"/seealso>
    class function GetTob(const Arguments, ArgumentID : String;
      const Separator: String = ';'; IgnoreCase: Boolean = True): Tob;

    ///<summary>
    ///compose un argument contenant l'adresse d'une Tob
    ///</summary>
    ///<param name="ArgumentID">nom de l'argument</param>
    ///<param name="ATob">La Tob à passer en argument</param>
    ///<param name="Separator">séparateur non pas d'arguments mais reliant ID & Valeur</param>
    ///<returns>ArgumentID+Separator+@ATob</returns>
    class function MakeFromTob(const ArgumentID: String; const ATob: Tob;
      const Separator: String = '='): String;
  end;

  ///<summary>
  ///Routine qui convertit un booléen à TRUE en caractère
  ///</summary>
  ///<returns>X</returns>
  ///<seealso cref="wFalse"/seealso>
  function wTrue: Char;
  ///<summary>
  ///Routine qui convertit un booléen à FALSE en caractère
  ///</summary>
  ///<returns>-</returns>
  ///<seealso cref="wTrue"/seealso>
  function wFalse: Char;

  ///<summary>
  ///Routine qui retourne le Carriage Return + Line Feed
  ///</summary>
  function CrLf: String;
  function CrLf_: WideString;

  ///<summary>
  ///Routine iif()
  ///</summary>
  ///<param name="Condition">Contient la partie évaluée</param>
  ///<param name="TruePart">Retour quand Condition = True</param>
  ///<param name="FalsePart">Retour quand Condition = False</param>
  ///<remarks>
  ///ATTENTION ne pas se faire avoir par la facilité ! 
  ///Les 2 parties retournées sont évaluées :
  ///Si on fait
  ///<example>
  ///  Result := iif(x < y, MonObjet_X.Propriété, MonObjet_Y.Propriété)
  ///</example>
  ///Si jamais <b>l'un ou l'autre</b> de MonObjet_X, MonObjet_Y vaut <b>nil</b>
  ///Access violation !
  ///</remarks>
  function iif(Condition, TruePart, FalsePart: Boolean): Boolean; overload;
  function iif(Condition: Boolean; const TruePart, FalsePart: Integer): Integer; overload;
  function iif(Condition: Boolean; const TruePart, FalsePart: Double): Double; overload;
  function iif(Condition: Boolean; const TruePart, FalsePart: String): String; overload;
  function iif(Condition: Boolean; const TruePart, FalsePart: HString): HString; overload;
  function iif(Condition: Boolean; const TruePart, FalsePart: Char): Char; overload;
  function iif(Condition: Boolean; const TruePart, FalsePart: Tob): Tob; overload;
  function iifV(Condition: Boolean; const TruePart, FalsePart: Variant): Variant;

  ///<summary>
  ///Routine qui permet de savoir si la valeur demandée est dans les limites imposées
  ///</summary>
  ///<param name="Value">Valeur à comparer</param>
  ///<param name="LowerValue">Valeur minimum</param>
  ///<param name="HigherValue">Valeur maximum</param>
  ///<param name="IncludingLower">Valeur mini à prendre en compte dans la limite</param>
  ///<param name="IncludingHigher">Valeur maxi à prendre en compte dans la limite</param>
  function isBetween(const Value, LowerValue, HigherValue: Double;
    IncludingLower: Boolean = True; IncludingHigher: Boolean = True): Boolean; overload;
  function isBetween(const Value, LowerValue, HigherValue: Integer;
    IncludingLower: Boolean = True; IncludingHigher: Boolean = True): Boolean; overload;

  ///<summary>
  ///routine qui permet de s'assurer que la division demandée est mathématiquement possible
  ///</summary>
  ///<param name="Dividande">Dividander</param>
  ///<param name="Divisor">Diviseur</param>
  ///<param name="Default">Valeur par défaut en cas d'erreur</param>
  ///<returns>Dividande / Divisor</returns>
  ///<remarks>
  ///en cas de division par 0, DivideDef <b>renvoie Default</b>
  ///</remarks>
 	function DivideDef(const Dividande, Divisor: Double; const Default: Double = 0): Double;

  ///<summary>
  ///routine qui permet de comparer 2 nombres flottants
  ///</summary>
  ///<param name="Value1">premier nombre</param>
  ///<param name="Value2">second nombre</param>
  ///<value>
  ///pour contourner un problème de delphi
  ///</value>
  function DoubleCompare(const Value1, Value2: Double): Boolean;

  ///<summary>
  ///Message d'erreur générique en cas d'erreur de transaction
  ///</summary>
  ///<value>
  ///Récupère un message d'erreur commun à tous pour signaler une interruption
  ///de traitement liée à une exception de transaction
  ///</value>
  function TransactionErrorMessage: WideString;

  procedure PutToDebugLog(const FunctionName: String; const Start: Boolean; const Msg: String = ''); //TS¤TODO deprecated

  ///<summary>
  ///permet de référencer l'interface ITypedGetter pour extraire des données
  ///d'une chaîne de caractères du type : ID1=Value1;ID2=Value2;IDn=ValueN...
  ///</summary>
  ///<param name="Arguments">chaîne à analyser</param>
  ///<param name="Separator">spérateur d'éléments de l'argument</param>
  ///<param name="IgnoreCase">tient compte ou non de la casse pour la recherche de l'ID</param>
  ///<returns>une référence à l'interface ITypedGetter</returns>
  function ArgumentReader(const Arguments: String;
    const Separator: String = ';'; IgnoreCase: Boolean = True): ITypedGetter;

implementation

uses
  SysUtils
  ,HCtrls
  ,gErrorConst
  ,cbpTrace
  ,DateUtils
  ,Math
  ,HZStream
  ;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function wTrue: char;
begin
  Result := BoolToStr_(True)[1];
end;
{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function wFalse: char;
begin
  Result := BoolToStr_(False)[1];
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function CrLf: String;
begin
  Result := #13#10
end;
{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function CrLf_: WideString;
begin
  Result := #13#10
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function iif(Condition, TruePart, FalsePart: Boolean): Boolean;
begin
	if Condition then
		Result := TruePart
	else
		Result := FalsePart
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function iif(Condition: Boolean; const TruePart, FalsePart: Integer): Integer;
begin
	if Condition then
		Result := TruePart
	else
		Result := FalsePart
end;
{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function iif(Condition: Boolean; const TruePart, FalsePart: Double): Double;
begin
	if Condition then
		Result := TruePart
	else
		Result := FalsePart
end;
{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function iif(Condition: Boolean; const TruePart, FalsePart: String): String;
begin
	if Condition then
		Result := TruePart
	else
		Result := FalsePart
end;
{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function iif(Condition: Boolean; const TruePart, FalsePart: HString): HString;
begin
	if Condition then
		Result := TruePart
	else
		Result := FalsePart
end;
{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function iif(Condition: Boolean; const TruePart, FalsePart: Char): Char;
begin   
	if Condition then
		Result := TruePart
	else
		Result := FalsePart
end;
{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function iif(Condition: Boolean; const TruePart, FalsePart: Tob): Tob;
begin
	if Condition then
		Result := TruePart
	else
		Result := FalsePart
end;
{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function iifV(Condition: Boolean; const TruePart, FalsePart: Variant): Variant;
begin
	if Condition then
		Result := TruePart
	else
		Result := FalsePart
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function isBetween(const Value, LowerValue, HigherValue: Double;
  IncludingLower: Boolean = True; IncludingHigher: Boolean = True): Boolean;
begin
  if          IncludingLower and     IncludingHigher then
    Result := (Value >= LowerValue) and (Value <= HigherValue)
  else if not IncludingLower and not IncludingHigher then
    Result := (Value >  LowerValue) and (Value <  HigherValue)
	else if     IncludingLower and not IncludingHigher then
    Result := (Value >= LowerValue) and (Value <  HigherValue)
  else if not IncludingLower and     IncludingHigher then
    Result := (Value >  LowerValue) and (Value <= HigherValue)
  else
    Result := False
end;
{ ------------------------------------------------------------------------------
------------------------------------------------------------------------------ }
function isBetween(const Value, LowerValue, HigherValue: Integer;
  IncludingLower: Boolean = True; IncludingHigher: Boolean = True): Boolean;
begin
  if          IncludingLower and     IncludingHigher then
    Result := (Value >= LowerValue) and (Value <= HigherValue)
  else if not IncludingLower and not IncludingHigher then
    Result := (Value >  LowerValue) and (Value <  HigherValue)
	else if     IncludingLower and not IncludingHigher then
    Result := (Value >= LowerValue) and (Value <  HigherValue)
  else if not IncludingLower and     IncludingHigher then
    Result := (Value >  LowerValue) and (Value <= HigherValue)
  else
    Result := False
end;

{ ------------------------------------------------------------------------------
------------------------------------------------------------------------------ }
function DivideDef(const Dividande, Divisor: Double; const Default: Double): Double;
begin
	if Divisor = 0 then
  	Result := Default
  else
   	Result := Dividande / Divisor
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function DoubleCompare(const Value1, Value2: Double): Boolean;
begin
  Result := ((Value1 = 0) and (Value2 = 0))
            or (Abs(Value1 - Value2) < Min(Abs(Value1), Abs(Value2)) * 1E-12)
end;

{ ------------------------------------------------------------------------------
------------------------------------------------------------------------------ }
function TransactionErrorMessage: WideString;
begin
  if V_Pgi.TransacErrorMessage <> '' then
  begin
    Result := V_Pgi.TransacErrorMessage;
    V_Pgi.TransacErrorMessage := ''
  end
  else
    Result := TraduireMemoire('Traitement interrompu en cours de transaction.')
            + CrLf_
            + TraduireMemoire('Veuillez relancer le traitement.')
end;

{ ------------------------------------------------------------------------------
------------------------------------------------------------------------------ }
procedure PutToDebugLog(const FunctionName: String; const Start: Boolean; const Msg: String = '');

    function _GetIoError: String;
    begin
      case V_Pgi.IoError of
        oeLettrage : Result := 'oeLettrage';
        oeOk       : Result := 'oeOk';
        oePiece    : Result := 'oePiece';
        oePointage : Result := 'oePointage';
        oeReseau   : Result := 'oeReseau';
        oeSaisie   : Result := 'oeSaisie';
        oeStock    : Result := 'oeStock';
        oeTiers    : Result := 'oeTiers';
      else
        {oeUnknown  : }
        Result := 'oeUnknown';
      end;
    end;

begin
  if V_PGI.Debug then
  begin
    Trace.TraceVerbose('Debug', '===========================================');
    Trace.TraceVerbose('Debug', ' Function: ' + FunctionName + iif(Start, ' >>>', ' <<<'));
    if Msg <> '' then
      Trace.TraceVerbose('Debug', ' Message : ' + Msg);
    Trace.TraceVerbose('Debug', ' IoError : ' + _GetIoError());
    Trace.TraceVerbose('Debug', '===========================================');
  end
end;





{ ******************************************************************************
  TGStrUtils
  ****************************************************************************** }

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.ArrayToStr(const Strings: ArrayOfStrings;
  const Separator: String): String;
var
  i: Integer;
begin
  Result := '';
  for i := Low(Strings) to High(Strings) do
    Result := Result + iif(Result <> '', Separator, '') + Strings[i]
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.ArrayToStr(const Values: ArrayOfVariants;
  const Separator: String): String;
var
  i: Integer;
begin
  Result := '';
  for i := Low(Values) to High(Values) do
    Result := Result + iif(Result <> '', Separator, '') + VarToStr(Values[i])
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.ArrayToStr(const Integers: ArrayOfIntegers;
  const Separator: String): String;
var
  i: Integer;
begin
  Result := '';
  for i := Low(Integers) to High(Integers) do
    Result := Result + iif(Result <> '', Separator, '') + IntToStr(Integers[i])
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.ArrayAddString(var Strings: ArrayOfStrings;
  const StrItem: String): Integer;
begin
  SetLength(Strings, Length(Strings) + 1);
  Result := High(Strings);
  Strings[Result] := StrItem
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.ArrayIndexOfString(const Strings: ArrayOfStrings;
  const StrItem: String): Integer;

  function _FindNext(const Index: Integer): Integer;
  begin
    if Index > High(Strings) then
      Result := -1
    else if Strings[Index] = StrItem then
      Result := Index
    else
      Result := _FindNext(Index + 1)
  end;

begin
  Result := _FindNext(Low(Strings))
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.CopyLeft(const Str: String;
  const Count: Integer): String;
begin
  Result := Copy(Str, 1, Count)
end;
class function TGStrUtils.CopyLeft(const Str: WideString;
  const Count: Integer): WideString;
begin
  Result := Copy(Str, 1, Count)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.CopyRight(const Str: String;
  const Count: Integer): String;
begin
	Result := Copy(Str, Length(Str) - Count + 1, Count)
end;
class function TGStrUtils.CopyRight(const Str: WideString;
  const Count: Integer): WideString;
begin
	Result := Copy(Str, Length(Str) - Count + 1, Count)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class procedure TGStrUtils.DeleteLeft(var Str: String; const Count: Integer);
begin
  Delete(Str, 1, Count)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class procedure TGStrUtils.DeleteRight(var Str: String;
  const Count: Integer);
begin
  Delete(Str, Length(Str) - Count + 1, Count)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.StringListToStr(const StringList: TStrings;
  const Separator: String): String;
var
  i: Integer;
begin
  if not Assigned(StringList) then
    raise EGSysObjectUnassigned.Create(
            G_CLASS_NOT_ASSIGNED,
            'gSysUtils.TGStrUtils.StringListToStr(StringList, Separator)',
            ['StringList', 'TStrings']
          );

  Result := '';
  for i := 0 to Pred(StringList.Count) do
    Result := Result + iif(Result <> '', Separator, '') + StringList[i]
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class procedure TGStrUtils.AsArray(Str: String;
  var Strings: ArrayOfStrings; const Separator: String);
begin
  { Attention ! cette procédure complète le tableau, elle ne le vide pas }
  while Str <> '' do
    ArrayAddString(Strings, Trim(ReadTokenPipe(Str, Separator)))
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.AsArray(const Str,
  Separator: string): ArrayOfStrings;
begin
  SetLength(Result, 0);
  AsArray(Str, Result, Separator)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class procedure TGStrUtils.StrToHStringList(Str: String;
  const Result: HTStrings; const Separator: String);
var
  Token: String;
begin
  if not Assigned(Result) then
    raise EGSysObjectUnassigned.Create(
            G_CLASS_NOT_ASSIGNED,
            'gSysUtils.TGStrUtils.StrToHStringList(s, Result, Separator)',
            ['Result', 'HTStrings']
          );
  while Str <> '' do
  begin
    Token := ReadTokenPipe(Str, Separator);
    Result.Add(Token)
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class procedure TGStrUtils.StrToStringList(Str: String;
  const Result: TStrings; const Separator: String);
var
  Token: String;
begin
  if not Assigned(Result) then
    raise EGSysObjectUnassigned.Create(
            G_CLASS_NOT_ASSIGNED,
            'gSysUtils.TGStrUtils.StrToStringList(s, Result, Separator)',
            ['Result', 'TStrings']
          );
  while Str <> '' do
  begin
    Token := ReadTokenPipe(Str, Separator);
    Result.Add(Token)
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.TokenCount(Str: String;
  const Separator: String): Integer;
begin
  Result := 0;
  while Str <> '' do
  begin
    ReadTokenPipe(Str, Separator);
    Inc(Result)
  end
end;


{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.TokenIndexOf(Str: String; const SubStr: String;
  const ExactSubStr: Boolean; const Separator: String): Integer;
var
  i: Integer;
  Str2: String;
begin
  Result := -1;
  i := 1;
  while (Str <> '') and (Result = -1) do
  begin
    Str2 := UpperCase(ReadTokenPipe(Str, Separator));
    if (ExactSubStr     and (UpperCase(SubStr) = Str2))
    or (not ExactSubStr and (Pos(UpperCase(SubStr), Str2) = 1)) then
      Result := i;
    Inc(i)
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.TokenReadFromIndex(Str: String;
  const Index: Integer; const Separator: String): String;
begin
  if (Str = '') or (Index <= 0) then
    Result := ''
  else if Index = 1 then
    Result := ReadTokenPipe(Str, Separator)
  else
  begin
    ReadTokenPipe(Str, Separator);
    Result := TokenReadFromIndex(Str, Index - 1, Separator)
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.TokenGetGreater(Str: String;
  const Separator: String = ';'): String;
var
  TempResult: String;
begin
  Result := '';
  TempResult := '';
  while Str <> '' do
  begin
    TempResult := ReadTokenPipe(Str, Separator);
    if Length(TempResult) > Length(Result) then
      Result := TempResult
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.ExtractParameters(const Str: String): String;
var
  i1,i2: integer;
begin
  Result := '';
  i1 := Pos('[', Str);
  if i1 > 0 then
  begin
  	i2 := Pos(']', Str);
   	if i2 > 0 then
    	Result := Copy(Str, i1 + 1, i2 - i1 - 1)
	end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.StringListCopyItem(const Item: String; FromList,
  ToList: HTStrings; ByItemName, ExcludeDoubloon,
  DeleteItemSource: Boolean; const ToIndex: Integer): Integer;
var
  FromIndexOf, ToIndexOf: Integer;
begin
  if ByItemName then
    Result := FromList.IndexOfName(Item)
  else
    Result := FromList.IndexOf(Item);
    
  FromIndexOf := Result;
  if Result >= 0 then
  begin
    if ByItemName then
      ToIndexOf := ToList.IndexOfName(Item)
    else
      ToIndexOf := ToList.IndexOf(Item);

    if ToIndexOf < 0 then
      Result := 0;

    if (Result = 0) or not ExcludeDoubloon then
    begin
      if ToIndex = -1 then
        ToList.Add(FromList[FromIndexOf])
      else
        ToList.Insert(ToIndex, FromList[FromIndexOf]);
      if DeleteItemSource then
        FromList.Delete(FromIndexOf)
    end
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.ExtractSection(const SectionName: String;
  var Buffer: String; const ModifyBuffer: Boolean): String;
var
  TS: TStringList;

  function _GetISection: Integer;
  var
    i: Integer;
  begin
    Result := -1;
    i := 0;
    while (Result = -1) and (i < TS.Count) do
    begin
      if Pos('[' + SectionName + ']', TS[i]) > 0 then
        Result := i;
      Inc(i)
    end
  end;

  function _EoSection(const i: Integer): Boolean;
  begin
    Result := (i >= TS.Count)
           or ((Pos('[', TS[i]) > 0) and (Pos('[', TS[i]) < Pos(']', TS[i])))
  end;

var
  TSSection: TStringList;
  i: Integer;
begin
  TS := TStringList.Create();
  try
    TSSection := TStringList.Create();
    try
      TS.SetText(PChar(Buffer));
      Result := '';
      i := _GetISection;
      if i >= 0 then
      begin
        TS[i] := Trim(StringReplace(TS[i], '[' + SectionName + ']', '', [rfIgnoreCase]));
        if TS[i] = '' then
          TS.Delete(i);
        while not _EoSection(i) do
        begin
          TSSection.Add(TS[i]);
          TS.Delete(i)
        end;
        Result := Trim(TSSection.Text);
        if ModifyBuffer then
          Buffer := TS.Text
      end
    finally
      TSSection.Free
    end
  finally
    TS.Free
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.PadLeft(const Str: String; const Count: Integer;
  const PadChar: Char): String;
begin
  if Length(Str) < Count then
    Result := StringOfChar(PadChar, Count - Length(Str)) + Str
  else
    Result := Copy(Str, 1, Count)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.PadRight(const Str: String;
  const Count: Integer; const PadChar: Char): String;
begin
  Result := CopyLeft(Str + StringOfChar(PadChar, Count), Count);
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.ExtractQuotes(const Str: String): String;
const
  Quote = '"';
begin
	Result := Str;
  if Length(Result) > 0 then
  begin
    if Result[1] = Quote then
      DeleteLeft(Result, 1);
    if Result[Length(Result)] = Quote then
      DeleteRight(Result, 1);
  end;
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.StringRepeat(const Str: string;
  Count: Integer): String;
begin
	Result := '';
  while Count > 0 do
  begin
   	Result := Result + Str;
    Dec(Count)
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.RTFToStr(
  const RTF: String;                       //> chaîne à convertir
  ReplaceLineFeedWithSpace: Boolean;       //> remplacer les CRLF par des ' '
  DoTrimLeft: Boolean;                     //> Supprimer les ' ' et CRLF de début
  const TrailAfter: Integer): String;      //> remplace tout ce qui est après ce nombre de caractère par '...'
var
  //n: Nombre de caractères à traiter à la source
  //x: Index du caractère que l'on traite dans la source (voir ThisChar)
  n, x: Integer;

  //Flag iniquant si on est en train de lier un code de formatage.
  //Un code de formatage commence par un "\"
  //et est considéré comme fini juste avant un autre "\", un " " ou un retour à la ligne.
  GettingCode: Boolean;  

  //Chaîne dans laquelle on stocke le code de formatage que l'on est en train de lire
  Code: String;

  //Caractère que l'on est en train de traiter, et caractère le précédant
  ThisChar, LastChar: Char; 

  //Niveau de groupe (ou bloc) de format dans lequel on se trouve 
  //un groupe commence par une accolade ouverte "{" et se termine par une accolade fermée "}"
  Group: Integer;

  //Flag indiquant si le caractère ThisChar doit être rejeté (True) ou recopié dans le résultat (False)
  Skip: Boolean;

  procedure _ProcessCode;
  begin
    { Si on vient de terminer la lecture d'un code de formatage }
    if ThisChar in ['\', ' ', #13, #10] then
    begin 
      { si on vient de lire le code de format d'un début de paragraphe
        ou celui d'un passage à la ligne... }
      if (Code = '\par') or (Code = '\line') or (Code = '\pard') then
      begin
        if ReplaceLineFeedWithSpace then
          Result := Result + ' '
        else
          Result := Result + CrLf;
        GettingCode := False;
        Skip := True
      end;

      { si on vient de lire le code de format d'une tabulation }
      if Code = '\tab' then
      begin 
        { #9 est le code iutilisé sous Windows pour coder une tabulation en code ASCII }
        Result := Result + #9;
        GettingCode := False;
        Skip := True
      end;
    end;

    { Un code de format \'xx indigue le code ASCII d'un caractère spécial
      (lettres accentuées en particulier). xx est ce code en héxadécimal. }
    if ((Length(Code) = 4) and ((Code[1] + Code[2])= '\''')) then
    begin 
      { on retranscrit le code hexadécimal en code ASCII) }
      Result := Result + Chr(StrToInt('$' + Code[3] + Code[4]));
      GettingCode := False;
      Skip := True
    end
  end;

begin 
  try
    { Initialisations }
    n := Length(RTF);
    Result := '';

    GettingCode := False;
    Group := 0;
    LastChar := #0;

    { Traitement de la source }
    x := 1;
    while x <= n do
    begin
      Skip := False;
      ThisChar := RTF[x];
      case ThisChar of
        '{': if LastChar <> '\' then
             begin
               { Début de groupe }
               Inc(Group);
               Skip := True;
             end
             else
               GettingCode := False;
        '}': if LastChar <> '\' then
             begin
               { Fin de groupe }
               Dec(Group);
               Skip := True;
             end
             else
               GettingCode := False;
        '\': if LastChar <> '\' then
             begin
               { Début de Code de format à traiter }
               if GettingCode then
                 _ProcessCode();
               Code := '';
               GettingCode := True;
             end
             else { c'était bien le caractère anti-slash (codé en RTF avec deux anti-slashs successifs }
               GettingCode := False;
        ' ': if GettingCode then
             begin
               { fin de Code de format }
               _ProcessCode();
               GettingCode := False;
               Skip := True;
             end;
        #10: begin
               if GettingCode then
               begin
                 { fin de Code de format }
                 _ProcessCode();
                 GettingCode := False;
                 Skip := True;
               end;
               { on est dans un groupe, on ne recopie pas le "LineFeed" }
               if Group > 0 then
                 Skip := True
             end;
        #13: begin
               if GettingCode then
               begin
                 _ProcessCode();
                 GettingCode := False;
                 Skip := True;
               end;
               { on est dans un groupe, on ne recopie pas le "Retour chariot" }
               if Group > 0 then
                 Skip := True;
             end;
      end;

      if not GettingCode then
      begin
        if not Skip and (Group <= 1) then
          { On a un caractère à recopier dans le résultat du traitement (du texte brut, pas du format) }
          Result := Result + ThisChar;
      end
      else
      begin
        { on lit le code }
        Code := Code + ThisChar;
        _ProcessCode();
      end;

      { Préparation de la boucle suivante }
      LastChar := ThisChar;
      Inc(x);
    end;
    { Fin du traitement de la source et début de traitement du résultat obtenu }
      
    { Suppression des catractères cr/lf et espaces en fin de chaîne }
    Result := TrimRight(Result);
    
    { idem mais pour le début de la chaîne }
    if DoTrimLeft then
      Result := TrimLeft(Result);

    { césure du texte si l'utilisateur l'à demandé }
    if TrailAfter > 0 then
    begin 
      if Length(Result) > TrailAfter then
      begin 
        SetLength(Result, TrailAfter);
        Result := Result + '...';
      end
    end
  except
    on E: Exception do
    begin
      Result := RTF;
      Trace.TraceVerbose('DEBUG', 'TGStrUtils.RTFToStr() : ' + E.Message);
    end
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.StrToAlphaNumStr(const Str,
  CharsToIgnore: String): String;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Length(Str) do
    if ((Str[i] >= 'a') and (Str[i] <= 'z'))
    or ((Str[i] >= 'A') and (Str[i] <= 'Z'))
    or ((Str[i] >= '0') and (Str[i] <= '9'))
    or (Pos(Str[i], CharsToIgnore) > 0) then
      Result := Result + Str[i]
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.Unzip(const Str: String): String;
begin
  if Str <> '' then
  try
    Result := DecompressAsciiString(Str)
  except
    Result := Str
  end
  else
    Result := Str
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGStrUtils.Zip(const Str: String): String;
begin
  if Str <> '' then
  try
    Result := CompressAsciiString(Str)
  except
    Result := Str
  end
  else
    Result := ''
end;





{ ******************************************************************************
  TGDateUtils
  ****************************************************************************** }

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.AddDay(const ADate: TDateTime;
  const Count: Integer; const HourReturned: TDayReturned): TDateTime;
var
  y, m, d, h, n, s, ms: Word;
begin
  case HourReturned of
    drSameDay : Result := PlusDate(ADate, Count, 'J');
    drFirstDay: begin
                  Result := PlusDate(ADate, Count, 'J');
                  DecodeDateTime(Result, y, m, d, h, n, s, ms);
                  Result := EncodeDateTime(y, m, d, 0, 0, 1, 0);
                end;
    drLastDay : begin
                  Result := PlusDate(ADate, Count, 'J');
                  DecodeDateTime(Result, y, m, d, h, n, s, ms);
                  Result := EncodeDateTime(y, m, d, 23, 59, 59, 99);
                end;
  else
    Result := ADate
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.AddMonth(const ADate: TDateTime;
  const Count: Integer; const DayReturned: TDayReturned): TDateTime;
var
  y, m, d, h, n, s, ms: Word;
begin
  case DayReturned of
    drSameDay : Result := PlusDate(ADate, Count, 'M');
    drFirstDay: begin
                  Result := PlusDate(ADate, Count, 'M');
                  DecodeDateTime(Result, y, m, d, h, n, s, ms);
                  Result := EncodeDateTime(y, m, 1, 0, 0, 1, 0);
                end;
    drLastDay : begin
                  Result  := PlusDate(ADate, Count + 1, 'M');
                  DecodeDateTime(Result, y, m, d, h, n, s, ms);
                  Result := FinDeMois(EncodeDateTime(y, m, 1, 23, 59, 59, 99) - 1)
                end;
  else
    Result := ADate
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.AddSemester(const ADate: TDateTime;
  const Count: Integer; const DayReturned: TDayReturned): TDateTime;
var
  y, m, d, h, n, s, ms: Word;
begin
  case DayReturned of
    drSameDay : Result := PlusDate(ADate, Count * 6, 'M');
    drFirstDay: begin
                  Result := PlusDate(ADate, Count * 6, 'M');
                  DecodeDateTime(Result, y, m, d, h, n, s, ms);
                  Result := EncodeDateTime(
                              y,
                              GetFirstMonthOfSemester(MonthToSemester(m)),
                              1,
                              0,
                              0,
                              1,
                              0
                            );
                end;
    drLastDay : begin
                  Result := PlusDate(ADate, (Count + 1) * 6, 'M');
                  DecodeDateTime(Result, y, m, d, h, n, s, ms);
                  Result := EncodeDateTime(
                              y,
                              GetFirstMonthOfSemester(MonthToSemester(m)),
                              1,
                              23,
                              59,
                              59,
                              99
                            ) - 1
                end;
  else
    Result := ADate
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.AddTrimester(const ADate: TDateTime;
  const Count: Integer; const DayReturned: TDayReturned): TDateTime;
var
  y, m, d, h, n, s, ms: Word;
begin
  case DayReturned of
    drSameDay : Result := PlusDate(ADate, Count * 3, 'M');
    drFirstDay: begin
                  Result  := PlusDate(ADate, Count * 3, 'M');
                  DecodeDateTime(Result, y, m, d, h, n, s, ms);
                  Result := EncodeDateTime(
                              y,
                              GetFirstMonthOfTrimester(MonthToTrimester(m)),
                              1,
                              0,
                              0,
                              1,
                              0
                            )
                end;
    drLastDay : begin
                  Result := PlusDate(ADate, (Count + 1) * 3, 'M');
                  DecodeDateTime(Result, y, m, d, h, n, s, ms);
                  Result := EncodeDateTime(
                              y,
                              GetFirstMonthOfTrimester(MonthToTrimester(m)),
                              1,
                              23,
                              59,
                              59,
                              99
                            ) - 1
                end;
  else
    Result := ADate
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.AddWeek(const ADate: TDateTime;
  const Count: Integer; const DayReturned: TDayReturned): TDateTime;
begin
  Result := ADate + 7 * Count;

  Case DayReturned of
    drFirstDay: while DayOfTheWeek(Result) > 1 do
                  Result := Result - 1;
    drLastDay : while DayOfTheWeek(Result) < 7 do
                  Result := Result + 1;
  end;
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.AddYear(const ADate: TDateTime;
  const Count: Integer; const DayReturned: TDayReturned): TDateTime;
var
  y, m, d, h, n, s, ms: Word;
begin
  case DayReturned of
    drSameDay : Result := PlusDate(ADate, Count, 'A');
    drFirstDay: begin
                  DecodeDateTime(ADate, y, m, d, h, n, s, ms);
                  Result := EncodeDateTime(y + Count, 1, 1, 0, 0, 1, 0);
                end;
    drLastDay : begin
                  DecodeDateTime(ADate, y, m, d, h, n, s, ms);
                  Result := EncodeDateTime(y + Count, 12, 31, 23, 59, 59, 99);
                end;
  else
    Result := ADate
  end;
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.FullDateFormatToStr(
  const AFormat: TFullDateFormat): String;
begin
  case AFormat of
    fdfYear     : Result := 'AAAA';
    fdfSemester : Result := 'AAAA-S';
    fdfTrimester: Result := 'AAAA-S-T';
    fdfMonth    : Result := 'AAAA-S-T-MM';
    fdfDay      : Result := 'AAAA-S-T-MM-JJ';
    fdfHour     : Result := 'AAAA-S-T-MM-JJ-HH';
    fdfMinute   : Result := 'AAAA-S-T-MM-JJ-HHMM';
    fdfSecond   : Result := 'AAAA-S-T-MM-JJ-HHMMSS';
  else
    raise EGSysDateFormatUnknown.Create(
            G_DATE_WRONGFORMAT,
            'gSysUtils.TGDateUtils.FullDateFormatToStr(AFormat)',
            ['AFormat', '???']
          )
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.StrToFullDateFormat(
  const StrFormat: String): TFullDateFormat;
begin
  if      StrFormat = 'AAAA' then
    Result := fdfYear
  else if StrFormat = 'AAAA-S' then
    Result := fdfSemester
  else if StrFormat = 'AAAA-S-T' then
    Result := fdfTrimester
  else if StrFormat = 'AAAA-S-T-MM' then
    Result := fdfMonth
  else if StrFormat = 'AAAA-S-T-MM-JJ' then
    Result := fdfDay
  else if StrFormat = 'AAAA-S-T-MM-JJ-HH' then
    Result := fdfHour
  else if StrFormat = 'AAAA-S-T-MM-JJ-HHMM' then
    Result := fdfMinute
  else if StrFormat = 'AAAA-S-T-MM-JJ-HHMMSS' then
    Result := fdfSecond
  else
    raise EGSysDateFormatUnknown.Create(
            G_DATE_WRONGFORMAT,
            'gSysUtils.TGDateUtils.StrToFullDateFormat(StrFormat)',
            ['StrFormat', StrFormat]
          )
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.FormatFullDate(const ADate: TDateTime;
  const AFormat: TFullDateFormat; const Truncate: Boolean): String;

  function _FormatDate(const NewDate: TDateTime): String;
  var
    y, m, d, h, n, s, ms: Word;
  begin
    DecodeDateTime(NewDate, y, m, d, h, n, s, ms);

    if not Truncate then
    begin
      case AFormat of
        fdfHour  : begin
                     n := 0;
                     s := 1;
                   end;
        fdfMinute: s := 1;
      end;
    end;

    Result := TGStrUtils.PadLeft(IntToStr(y), 4, '0')                   //AAAA (peu probable les 0xxx mais bon...)
            + '-' + FloatToStr(TGRounding.Round(m / 6, rp1, rndHigher)) //-S
            + '-' + FloatToStr(TGRounding.Round(m / 3, rp1, rndHigher)) //-T
            + '-' + TGStrUtils.PadLeft(IntToStr(m), 2, '0')             //-MM
            + '-' + TGStrUtils.PadLeft(IntToStr(d), 2, '0')             //-JJ
            + '-' + TGStrUtils.PadLeft(IntToStr(h), 2, '0')             //-HH
            +       TGStrUtils.PadLeft(IntToStr(n), 2, '0')             //MM
            +       TGStrUtils.PadLeft(IntToStr(s), 2, '0')             //SS
  end;

var
  NewDate : TDateTime;
begin
  if Truncate then
  begin
    Result := _FormatDate(ADate);
    Result := TGStrUtils.CopyLeft(Result, Length(FullDateFormatToStr(AFormat)));
  end
  else
  begin
    case AFormat of
      fdfYear     : NewDate := AddYear     (ADate, 0, drFirstDay);
      fdfSemester : NewDate := AddSemester (ADate, 0, drFirstDay);
      fdfTrimester: NewDate := AddTrimester(ADate, 0, drFirstDay);
      fdfMonth    : NewDate := AddMonth    (ADate, 0, drFirstDay);
      fdfDay      : NewDate := AddDay      (ADate, 0, drFirstDay);
    else
      NewDate := ADate;
    end;

	  Result := _FormatDate(NewDate)
  end;
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.GetYear(const ADate: TDateTime): Word;
var
  m, d: Word;
begin
  DecodeDate(ADate, Result, m, d)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.GetMonth(const ADate: TDateTime): Word;
var
  y, d: Word;
begin
  DecodeDate(ADate, y, Result, d)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.GetDay(const ADate: TDateTime): Word;
var
  y, m: Word;
begin
  DecodeDate(ADate, y, m, Result)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.GetHour(const ATime: TTime): Word;
var
  n, s, ms: Word;
begin
  DecodeTime(ATime, Result, n, s, ms);
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.GetMinute(const ATime: TTime): Word;
var
  h, s, ms: Word;
begin
  DecodeTime(ATime, h, Result, s, ms);
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.GetSecond(const ATime: TTime): Word;
var
  h, n, ms: Word;
begin
  DecodeTime(ATime, h, n, Result, ms);
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.GetMilliSecond(const ATime: TTime): Word;
var
  h, n, s: Word;
begin
  DecodeTime(ATime, h, n, s, Result);
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.GetDayNumber(const ADate: TDateTime): Integer;
begin
  Result := Trunc(ADate - EncodeDate(GetYear(ADate), 1, 1)) + 1
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.MonthToStr(const Month: Word;
  const UpCase1rst: Boolean): HString;
begin
  case Month of
    1 : Result := TraduireMemoire('janvier');
    2 : Result := TraduireMemoire('février');
    3 : Result := TraduireMemoire('mars');
    4 : Result := TraduireMemoire('avril');
    5 : Result := TraduireMemoire('mai');
    6 : Result := TraduireMemoire('juin');
    7 : Result := TraduireMemoire('juillet');
    8 : Result := TraduireMemoire('août');
    9 : Result := TraduireMemoire('septembre');
    10: Result := TraduireMemoire('octobre');
    11: Result := TraduireMemoire('novembre');
    12: Result := TraduireMemoire('décembre');
  else
    Result := '';
  end;

  if UpCase1rst and (Result <> '') then
  	Result[1] := UpCase_(Result[1]);
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.GetLastWeek(const ADate: TDateTime): Word;
var
	y, m, d: Word;
begin
	DecodeDate(ADate, y, m, d);
  d := 31;
  m := 12;
  Result := NumSemaine(EncodeDate(y, m, d));
  while Result = 1 do
  begin
  	Dec(d);
		Result := NumSemaine(EncodeDate(y, m, d));
  end;
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.MonthToTrimester(const Month: Word): Word;
begin
  Assert(
    isBetween(Month, 1, 12),
    Format(
      'gSysUtils.MonthToTrimester(Month) > Indice de mois hors limites (%d)',
      [Month]
    )
  );
  Result := (Month + 2) div 3
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.GetTrimester(const ADate: TDateTime): Word;
begin
  Result := MonthToTrimester(GetMonth(ADate))
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.GetFirstMonthOfTrimester(
  const Trimester: Word): Word;
begin
  Assert(
    isBetween(Trimester, 1, 4),
    Format(
      'gSysUtils.GetFirstMonthOfTrimester(Trimester) > Indice de trimestre hors limites (%d)',
      [Trimester]
    )
  );
  Result := Trimester * 3 - 2
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.MonthToSemester(const Month: Word): Word;
begin
  Assert(
    isBetween(Month, 1, 12),
    Format(
      'gSysUtils.MonthToSemester(Month) > Indice de mois hors limites (%d)',
      [Month]
    )
  );
  Result := (Month + 5) div 6
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.GetFirstMonthOfSemester(
  const Semester: Word): Word;
begin
  Assert(
    isBetween(Semester, 1, 2),
    Format(
      'gSysUtils.GetFirstMonthOfSemester(Semester) > Indice de semester hors limites (%d)',
      [Semester]
    )
  );
  Result := Semester * 6 - 5
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.GetSemester(const ADate: TDateTime): Word;
begin
  Result := MonthToSemester(GetMonth(ADate))
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.isInvalidDate(const ADate: String): Boolean;
begin
  Result := False;
  try
    StrToDate(ADate)
  except
    on E: EConvertError do
      Result := True
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.TimeToHHCC(const ATime: Double): Double;
var
  nDec: Extended;
	sDec: String;
	Err, hh, mm, ss, pp: Integer;
  Negative : Boolean;
begin
  Result := 0;
  Err := 0;
  Negative := ATime < 0;
  if Abs(ATime) < 999999 then
  begin
    { Extrait la partie entière de ATime }
    hh := Trunc(Abs(ATime));
    { Formate et Extrait la partie décimale de ATime }
    nDec := Frac(Abs(ATime));
    sDec := Format('%8.6f', [nDec]);
    { Extrait les valeurs dans la chaine de la partie décimale formatée }
    mm := ValeurI(Copy(sDec, 3, 2));
    ss := ValeurI(Copy(sDec, 5, 2));
    pp := ValeurI(Copy(sDec, 7, 2));
    if (mm >= 60) or (ss >= 60) then
      Err := 2
    else
      Result := Arrondi(hh + (mm / 60) + (ss / 3600) + (pp / 360000), 6);
  end
  else
    Err := 1;
  if Negative and (Result <> 0) then
    Result := - Result;

  Assert(
    Err = 0,
    Format(
      'Conversion de temps impossible (TGDateUtils.TimeToHHCC). %s n''est pas un temps correct !',
      [FloatToStr(ATime)]
    )
  );
  if Err <> 0 then
    Result := 0
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGDateUtils.HHCCToTime(const ATime: Double;
  const Format: String): Double;
var
  TpsSS: Double;
  hh, mn, ss: Integer;
  Negative: Boolean;
begin
  Negative := ATime < 0;

  mn := 0;
  ss := 0;

  { Les heures }
  hh := Trunc(Abs(ATime));
  if Format <> 'HH' then
  begin
    { Les secondes totales}
    TpsSS := (Abs(ATime) - hh) * 3600;

    { Les minutes }
    mn := Max(0, Min(59, Trunc(TpsSS / 60)));

    { Les secondes }
    if Format <> 'MM' then
    begin
      if      Format = 'SS' then
        ss := Max(0, Min(59, Round(TpsSS - mn * 60)))
      else if Format = 'MS' then
        ss := Max(0, Min(59, Trunc(TpsSS - mn * 60)))
      else
        ss := Max(0, Min(59, Round(TpsSS - mn * 60)));
    end
  end;

  { Résultat }
  Result := hh + mn/100 + ss/10000;

  if Format = 'MS' then
    Result := Result + ((Abs(ATime) - hh) * 3600 - Trunc((Abs(ATime) - hh) * 3600)) / 10000;

  if Negative then
    Result := - Result
end;





{ ******************************************************************************
  TGRounding
  ****************************************************************************** }

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGRounding.NbDecimalToRoundingPrecision(
  const NbDecimal: Integer): TRoundingPrecision;
begin
  case NbDecimal of
    -5: Result := rp100000;
    -4: Result := rp10000;
    -3: Result := rp1000;
    -2: Result := rp100;
    -1: Result := rp10;
     0: Result := rp1;
    +1: Result := rp01;
    +2: Result := rp001;
    +3: Result := rp0001;
    +4: Result := rp00001;
  else
    Result := rp00001;
    if Trace.IsTraceWarningEnabled() then
      TraceWarning(Format(G_ROUNDING_NBDECIMALOFFBOUNDS, [NbDecimal]));
    if V_Pgi.Sav then
      raise EGSysRoundingNbdecimalOffBounds.Create(
              G_ROUNDING_NBDECIMALOFFBOUNDS,
              'gSysUtils.TGRounding.NbDecimalToRoundingPrecision(NbDecimal)'
              [NbDecimal]
            )
  end;
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGRounding.Round(const Value: Double;
  const RoundingPrecision: TRoundingPrecision;
  const RoundingMethod: TRoundingMethod): Double;
var
  Weight: Double;
begin
  Weight := RoundingPrecisionToWeight(RoundingPrecision);

  Result := Value / Weight;
  case RoundingMethod of
    rndLower  : Result := Trunc(Result);
    rndNearest: Result := Arrondi(Result, 0);
    rndHigher : Result := Arrondi(Result + 0.4989, 0);
  end;
  Result := Result * Weight
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGRounding.RoundByMultipleOf(const Value,
  MultipleOf: Double): Double;
var
  Nb, NbSup: Double;
begin
  if MultipleOf = 0 then
    Result := Value
  else
  begin
    Nb    := Value / MultipleOf;
    NbSup := Self.Round(Nb, rp1, rndHigher);
    if Nb <> NbSup then
      Result := NbSup * MultipleOf
    else
      Result := Value
  end;
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGRounding.RoundingPrecisionToNbDecimal(
  const RoundingPrecision: TRoundingPrecision): Integer;
begin
  case RoundingPrecision of
    rp100000 : Result := -5;
    rp10000  : Result := -4;
    rp1000   : Result := -3;
    rp100    : Result := -2;
    rp10     : Result := -1;
    rp1      : Result :=  0;
    rp01     : Result := +1;
    rp001    : Result := +2;
    rp0001   : Result := +3;
    rp00001  : Result := +4;
  else
    Result := 0;
    if Trace.IsTraceWarningEnabled() then
      TraceWarning(G_ROUNDING_PRECISIONOFFBOUNDS);
    if V_Pgi.Sav then
      raise EGSysRoundingPrecisionOffBounds.Create(
              G_ROUNDING_PRECISIONOFFBOUNDS,
              'gSysUtils.TGRounding.RoundingPrecisionToNbDecimal(RoundingPrecision)'
            )
  end;
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGRounding.RoundingPrecisionToWeight(
  const RoundingPrecision: TRoundingPrecision): Double;
begin
  case RoundingPrecision of
    rp100000 : Result := 100000     ;
    rp10000  : Result :=  10000     ;
    rp1000   : Result :=   1000     ;
    rp100    : Result :=    100     ;
    rp10     : Result :=     10     ;
    rp1      : Result :=      1     ;
    rp01     : Result :=      0.1   ;
    rp001    : Result :=      0.01  ;
    rp0001   : Result :=      0.001 ;
    rp00001  : Result :=      0.0001;
  else
    Result := 0.0;
    if Trace.IsTraceWarningEnabled() then
      TraceWarning(G_ROUNDING_PRECISIONOFFBOUNDS);
    if V_Pgi.Sav then
      raise EGSysRoundingPrecisionOffBounds.Create(
              G_ROUNDING_PRECISIONOFFBOUNDS,
              'gSysUtils.TGRounding.RoundingPrecisionToWeight(RoundingPrecision)'
            )
  end;
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGRounding.StrToRoundingContext(const Str: String): TRoundingContext;
begin
  if      Str = 'T' then Result := rcTaux
  else if Str = 'P' then Result := rcPrix
  else if Str = 'M' then Result := rcMont
  else //mode secure, voire à décliener en StrToroundingContextDef()
  begin
    Result := rcMont;
    if Trace.IsTraceWarningEnabled() then
      TraceWarning(Format(G_ROUNDING_CONTEXTUNKNOWN, [Str]));
    if V_Pgi.Sav then
      raise EGSysRoundingContextUnknown.Create(
              G_ROUNDING_CONTEXTUNKNOWN,
              'gSysUtils.TGRounding.StrToRoundingContext(Str)',
              [Str]
            )
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGRounding.RoundingContextToStr(
  const RoundingContext: TRoundingContext): String;
begin
  case RoundingContext of
    rcTaux: Result := 'T';
    rcPrix: Result := 'P';
    rcMont: Result := 'M';
  else
    Result := 'M';
    if Trace.IsTraceWarningEnabled() then
      TraceWarning(Format(G_ROUNDING_CONTEXTUNKNOWN, [IntToStr(Integer(RoundingContext))]));
    if V_Pgi.Sav then
      raise EGSysRoundingContextUnknown.Create(
              G_ROUNDING_CONTEXTUNKNOWN,
              'gSysUtils.TGRounding.RoundingContextToStr(RoundingContext)',
              [IntToStr(Integer(RoundingContext))]
            )
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGRounding.StrToRoundingMethod(const Str: String): TRoundingMethod;
begin
  if      Str = 'I' then Result := rndLower
  else if Str = 'P' then Result := rndNearest
  else if Str = 'S' then Result := rndHigher
  else //mode safe, voire à décliner en StrToRoundingMethodDef()
  begin
    Result := rndNearest;
    if Trace.IsTraceWarningEnabled() then
      TraceWarning(Format(G_ROUNDING_METHODUNKNOWN, [Str]));
    if V_Pgi.Sav then
      raise EGSysRoundingMethodUnknown.Create(
              G_ROUNDING_METHODUNKNOWN,
              'gSysUtils.TGRounding.StrToRoundingMethod(Str)',
              [Str]
            )
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGRounding.RoundingMethodToStr(const RoundingMethod: TRoundingMethod): String;
begin
  case RoundingMethod of
    rndLower  : Result := 'I';
    rndNearest: Result := 'P';
    rndHigher : Result := 'S';
  else
    Result := 'P';
    if Trace.IsTraceWarningEnabled() then
      TraceWarning(Format(G_ROUNDING_METHODUNKNOWN, [IntToStr(Integer(RoundingMethod))]));
    if V_Pgi.Sav then
      raise EGSysRoundingMethodUnknown.Create(
              G_ROUNDING_METHODUNKNOWN,
              'gSysUtils.TGRounding.RoundingMethodToStr(RoundingMethod)',
              [IntToStr(Integer(RoundingMethod))]
            )
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGRounding.WeightToRoundingPrecision(
  const Weight: Double): TRoundingPrecision;
begin
  if      DoubleCompare(Weight, 100000     ) then Result := rp100000
  else if DoubleCompare(Weight,  10000     ) then Result := rp10000
  else if DoubleCompare(Weight,   1000     ) then Result := rp1000
  else if DoubleCompare(Weight,    100     ) then Result := rp100
  else if DoubleCompare(Weight,     10     ) then Result := rp10
  else if DoubleCompare(Weight,      1     ) then Result := rp1
  else if DoubleCompare(Weight,      0.1   ) then Result := rp01
  else if DoubleCompare(Weight,      0.01  ) then Result := rp001
  else if DoubleCompare(Weight,      0.001 ) then Result := rp0001
  else if DoubleCompare(Weight,      0.0001) then Result := rp00001
  else
  begin
    Result := rp00001;
    if Trace.IsTraceWarningEnabled() then
      TraceWarning(Format(G_ROUNDING_WEIGHTOFFBOUNDS, [Weight]));
    if V_Pgi.Sav then
      raise EGSysRoundingWeightOffBounds.Create(
              G_ROUNDING_WEIGHTOFFBOUNDS,
              'gSysUtils.TGRounding.WeightToRoundingPrecision(Weight)',
              [Weight]
            )
  end;
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class procedure TGRounding.TraceWarning(const Msg: String);
const
  RND_TRACE_CATEGORY = 'G_ROUNDING';
begin
  Trace.TraceWarning(RND_TRACE_CATEGORY, Msg)
end;

class function TGRounding.ArrondiMeth2RoundingMethod(
  const ArrondiMeth: tArrondiMeth): TRoundingMethod;
begin
  case ArrondiMeth of
    tAMInferieure: Result := rndLower;
    tAMSuperieure: Result := rndHigher;
    tAMPlusProche: Result := rndNearest;
  else
    raise EConvertError.Create('Erreur de conversion : ArrondiMeth2RoundingMethod()')
  end;
end;

class function TGRounding.ArrondiPrec2RoundingPrecision(
  const ArrondiPrec: tArrondiPrec): TRoundingPrecision;
begin
  case ArrondiPrec of
    tAPCentaineDeMillier: Result := rp100000;
    tAPDizaineDeMillier : Result := rp10000;
    tAPMillier          : Result := rp1000;
    tAPCentaine         : Result := rp100;
    tAPDizaine          : Result := rp10;
    tAPUnite            : Result := rp1;
    tAPDizieme          : Result := rp01;
    tAPCentieme         : Result := rp001;
    tAPMillieme         : Result := rp0001;
    tAPDixMillieme      : Result := rp00001;
  else
    raise EConvertError.Create('Erreur de conversion : ArrondiPrec2RoundingPrecision()')
  end;
end;

class function TGRounding.ArrondiQuoi2RoundingContext(
  const ArrondiQuoi: tArrondiQuoi): TRoundingContext;
begin
  case ArrondiQuoi of
    tTPMTaux: Result := rcTaux;
    tTPMPrix: Result := rcPrix;
    tTPMMont: Result := rcMont;
  else
    raise EConvertError.Create('Erreur de conversion : ArrondiQuoi2RoundingContext()')
  end;
end;




{ ******************************************************************************
  TGComUtils
  ****************************************************************************** }

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGComUtils.IntToC2(const Value: Integer): String;
begin
  Assert(
    isBetween(Value, 0, 99),
    Format('gSysUtils.TGComUtils.IntToC2(Value) > Valeur hors limites (%d)', [Value])
  );

  Result := TGStrUtils.PadLeft(IntToStr(Value), 2, '0')
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGComUtils.isVarInt(const Value: Variant): Boolean;
begin
  Result := VarType(Value) in [varShortInt, varWord, varByte, varLongWord, varInt64, varInteger]
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGComUtils.GetArrayOfStrings(
  const ArrayData: array of String): ArrayOfStrings;
var
  i: Integer;
begin
  SetLength(Result, Length(ArrayData));
  for i := Low(ArrayData) to High(ArrayData) do
    Result[i] := ArrayData[i]
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGComUtils.GetArrayOfVariants(
  const ArrayData: array of Variant): ArrayOfVariants;
var
  i: Integer;
begin
  SetLength(Result, Length(ArrayData));
  for i := Low(ArrayData) to High(ArrayData) do
    Result[i] := ArrayData[i]
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGComUtils.StrToAction(const Action: String): TActionFiche;
begin
  if Action = 'CONSULTATION' then
    Result := taConsult
  else if Action = 'CREATION' then
    Result := taCreat
  else
    Result := taModif
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGComUtils.ActionToStr(const Action: TActionFiche): String;
begin
  case Action of
    taCreat, taCreatEnSerie, taCreatOne: Result := 'CREATION';
    taModif, taModifEnSerie            : Result := 'MODIFICATION';
    taDuplique                         : Result := 'DUPLICATION';
  else
    Result := 'CONSULTATION'
  end
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGComUtils.JAiLeDroitConsult(const Prefix: String): Boolean;
begin
  if (CtxAffaire in V_PGI.PgiContexte) then
  begin
    if      Prefix = 'GA'  then Result := JAiLeDroitTag( 71511)
    else if Prefix = 'TAR' then Result := JAiLeDroitTag( 74346)
    else if Prefix = 'FN1' then Result := JAiLeDroitTag( 74341)
    else if Prefix = 'FN2' then Result := JAiLeDroitTag( 74342)
    else if Prefix = 'FN3' then Result := JAiLeDroitTag( 74343)
    else if Prefix = 'YTP' then Result := JAiLeDroitTag(215102)
    else if Prefix = 'ARS' then Result := JAiLeDroitTag( 71104)
    else if Prefix = 'T'   then Result := JAiLeDroitTag( 71101)
    else if Prefix = 'TRC' then Result := JAiLeDroitTag( 74315)
    else if Prefix = 'TRF' then Result := JAiLeDroitTag( 74318)
    else
    begin
      Result := False;
      if V_Pgi.Sav then
        raise EGSysFeatureTagUnknown.Create(
                G_RIGHTS_USERACCESS,
                'gCtrlUtils.TGCtrlsUtils.JAiLeDroitConsult(Prefix)',
                [Prefix]
              )
    end;
  end
  else
  begin
    if      Prefix = 'WNT' then Result := JAiLeDroitTag(126220)
    else if Prefix = 'WGT' then Result := JAiLeDroitTag(126230)
    else if Prefix = 'WPC' then Result := JAiLeDroitTag(125210)
    else if Prefix = 'WOT' then Result := JAiLeDroitTag(120210)
    else if Prefix = 'WOL' then Result := JAiLeDroitTag(120220)
    else if Prefix = 'WOP' then Result := JAiLeDroitTag(120230)
    else if Prefix = 'WOB' then Result := JAiLeDroitTag(120240)
    else if Prefix = 'WPH' then Result := JAiLeDroitTag(120022)
    else if Prefix = 'QIT' then Result := JAiLeDroitTag(121210)
    else if Prefix = 'QWB' then Result := JAiLeDroitTag(352210)
    else if Prefix = 'QCI' then Result := JAiLeDroitTag(121220)
//DKZ_SCRUM01_AccesAtelierDepuisFicheDetailCircuit
    else if Prefix = 'QSI' then Result := JAiLeDroitTag(121630) or JAiLeDroitTag(127122)
    else if Prefix = 'GA'  then Result := JAiLeDroitTag( 32501)
    else if Prefix = 'GCA' then Result := JAiLeDroitTag( 31702)
    else if Prefix = 'TAR' then Result := JAiLeDroitTag( 65302)
    else if Prefix = 'FN1' then Result := JAiLeDroitTag( 65321)
    else if Prefix = 'FN2' then Result := JAiLeDroitTag( 65322)
    else if Prefix = 'FN3' then Result := JAiLeDroitTag( 65323)
    else if Prefix = 'GM'  then Result := JAiLeDroitTag(211901)
    else if Prefix = 'GCQ' then Result := JAiLeDroitTag(211911)
    else if Prefix = 'FVS' then Result := JAiLeDroitTag(212301)
    else if Prefix = 'YTP' then Result := JAiLeDroitTag(215102)
    else if Prefix = 'ARS' then Result := JAiLeDroitTag(126240)
    else if Prefix = 'T'   then Result := JAiLeDroitTag( 30501)
    else if Prefix = 'TRC' then Result := JAiLeDroitTag( 65108)
    else if Prefix = 'TRF' then Result := JAiLeDroitTag( 65109)
    else if Prefix = 'SCC' then Result := JAiLeDroitTag( 65106)
    else if Prefix = 'GOR' then Result := JAiLeDroitTag( 65107)
    else if Prefix = 'REN' then Result := JAiLeDroitTag( 65110)
    else if Prefix = 'WXC' then Result := JAiLeDroitTag(126150)
    else if Prefix = 'WXT' then Result := JAiLeDroitTag(126150)
    else if Prefix = 'GVP' then Result := JAiLeDroitTag(212303)
    else if Prefix = 'GCL' then Result := JAiLeDroitTag( 30503)
    else if Prefix = 'GCV' then Result := JAiLeDroitTag( 30508)
    else
    begin
      Result := False;
      if V_Pgi.Sav then
        raise EGSysFeatureTagUnknown.Create(
                G_RIGHTS_USERACCESS,
                'gCtrlUtils.TGCtrlsUtils.JAiLeDroitConsult(Prefix)',
                [Prefix]
              )
    end;
  end;
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGComUtils.JAiLeDroitGestion(const Prefix: String): Boolean;
begin
  if (CtxAffaire in V_PGI.PgiContexte) then
  begin
    if      Prefix = 'GA'      then Result := JAiLeDroitTag( 71511)
    else if Prefix = 'TAR'     then Result := JAiLeDroitTag( 74346)
    else if Prefix = 'FN1'     then Result := JAiLeDroitTag( 74341)
    else if Prefix = 'FN2'     then Result := JAiLeDroitTag( 74342)
    else if Prefix = 'FN3'     then Result := JAiLeDroitTag( 74343)
    else if Prefix = 'YTP'     then Result := JAiLeDroitTag(215102)
    else if Prefix = 'ARS'     then Result := JAiLeDroitTag( 71104)
    else if Prefix = 'T'       then Result := JAiLeDroitTag( 71101)
    else if Prefix = 'TRC'     then Result := JAiLeDroitTag( 74315)
    else if Prefix = 'TRF'     then Result := JAiLeDroitTag( 74318)
    else if Prefix = 'GCA'     then Result := JAiLeDroitTag(138602)
    else
    begin
      Result := False;
      if V_Pgi.Sav then
        raise EGSysFeatureTagUnknown.Create(
                G_RIGHTS_USERACCESS,
                'gCtrlUtils.TGCtrlsUtils.JAiLeDroitGestion(Prefix)',
                [Prefix]
              )
    end;
  end
  else
  begin
    if      Prefix = 'WNT' then Result := JAiLeDroitTag(126120)
    else if Prefix = 'WGT' then Result := JAiLeDroitTag(126130)
    else if Prefix = 'WAN' then Result := JAiLeDroitTag(126150)
    else if Prefix = 'WPE' then Result := JAiLeDroitTag(126420)
    else if Prefix = 'WPC' then Result := JAiLeDroitTag(125110)
    else if Prefix = 'WOT' then Result := {$IFDEF GPAO}JAiLeDroitTag(120110){$ELSE GPAO}JAiLeDroitTag(267001){$ENDIF GPAO} //TS¤ToDo pas de ifdef
    else if Prefix = 'WOL' then Result := {$IFDEF GPAO}JAiLeDroitTag(120120){$ELSE GPAO}JAiLeDroitTag(267002){$ENDIF GPAO} //TS¤ToDo pas de ifdef
    else if Prefix = 'WOP' then Result := JAiLeDroitTag(120130)
    else if Prefix = 'WOB' then Result := {$IFDEF GPAO}JAiLeDroitTag(120140){$ELSE GPAO}JAiLeDroitTag(267003){$ENDIF GPAO} //TS¤ToDo pas de ifdef
    else if Prefix = 'WLB' then Result := JAiLeDroitTag(128120)
    else if Prefix = 'WRB' then Result := JAiLeDroitTag(128130)
    else if Prefix = 'QWB' then Result := JAiLeDroitTag(352110)
    else if Prefix = 'QIT' then Result := JAiLeDroitTag(121110)
    else if Prefix = 'QCI' then Result := JAiLeDroitTag(121120)
//DKZ_SCRUM01_AccesAtelierDepuisFicheDetailCircuit
    else if Prefix = 'QSI' then Result := JAiLeDroitTag(121630) or JAiLeDroitTag(127122)
    else if Prefix = 'GA'  then Result := JAiLeDroitTag(325011)
    else if Prefix = 'GCA' then Result := JAiLeDroitTag(31702)
    else if Prefix = 'TAR' then Result := JAiLeDroitTag( 65302)
    else if Prefix = 'FN1' then Result := JAiLeDroitTag( 65321)
    else if Prefix = 'FN2' then Result := JAiLeDroitTag( 65322)
    else if Prefix = 'FN3' then Result := JAiLeDroitTag( 65323)
    else if Prefix = 'GM'  then Result := JAiLeDroitTag(211901)
    else if Prefix = 'GCQ' then Result := JAiLeDroitTag(211911)
    else if Prefix = 'FVS' then Result := JAiLeDroitTag(212301)
    else if Prefix = 'YTP' then Result := JAiLeDroitTag(215102)
    else if Prefix = 'ARS' then Result := JAiLeDroitTag(126314)
    else if Prefix = 'T'   then Result := JAiLeDroitTag( 30501)
    else if Prefix = 'TRC' then Result := JAiLeDroitTag( 65108)
    else if Prefix = 'TRF' then Result := JAiLeDroitTag( 65109)
    else if Prefix = 'WWF' then Result := JAiLeDroitTag(215205)
    else if Prefix = 'RQN' then Result := JAiLeDroitTag( 124110)
    else if Prefix = 'RQD' then Result := JAiLeDroitTag( 124120)
    else if Prefix = 'RQP' then Result := JAiLeDroitTag( 124130)
    else if Prefix = 'RQT' then Result := JAiLeDroitTag( 124742)
    else if Prefix = 'RAC' then Result := JAiLeDroitTag( 92104)
    else if Prefix = 'SCC' then Result := JAiLeDroitTag( 65106)
    else if Prefix = 'GOR' then Result := JAiLeDroitTag( 65107)
    else if Prefix = 'REN' then Result := JAiLeDroitTag( 65110)
    else if Prefix = 'WXC' then Result := JAiLeDroitTag(126150)
    else if Prefix = 'WXT' then Result := JAiLeDroitTag(126150)
    else if Prefix = 'GVP' then Result := JAiLeDroitTag(212303)
    else if Prefix = 'GCL' then Result := JAiLeDroitTag( 30503)
    else if Prefix = 'GCV' then Result := JAiLeDroitTag( 30508)
    else if Prefix = 'PRO' then Result := JAiLeDroitTag(128110)
    else if Prefix = 'AFF' then Result := JAiLeDroitTag(128120)
    {$IFDEF BIBS}  //TS¤ToDo pas de ifdef
    else if Prefix = 'E01' then Result := JAiLeDroitTag(305110)
    {$ENDIF BIBS}
    else
    begin
      Result := False;
      if V_Pgi.Sav then
        raise EGSysFeatureTagUnknown.Create(
                G_RIGHTS_USERACCESS,
                'gCtrlUtils.TGCtrlsUtils.JAiLeDroitGestion(Prefix)',
                [Prefix]
              )
    end;
  end;
end;






{ ******************************************************************************
  TGArgumentGetter
  ****************************************************************************** }

type
  TGArgumentGetter = class(TInterfacedObject, ITypedGetter)
  private
    FArguments: String;
    FSeparator: String;
    FIgnoreCase: Boolean;
  public
    function GetValue(const Id: String): Variant;
    function GetString(const Id: String): String;
    function GetInteger(const Id: String): Integer;
    function GetDouble(const Id: String): Double;
    function GetBoolean(const Id: String): Boolean;
    function GetDateTime(const Id: String): TDateTime;
    function GetObject(const Id: String): TObject;

    constructor Create(const AArguments, ASeparator: String; AIgnoreCase: Boolean);

    property Arguments: String read FArguments;
    property Separator: String read FSeparator;
    property IgnoreCase: Boolean read FIgnoreCase;
  end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function ArgumentReader(const Arguments: String; const Separator: String = ';';
  IgnoreCase: Boolean = True): ITypedGetter;
begin
  Result := TGArgumentGetter.Create(Arguments, Separator, IgnoreCase) as ITypedGetter
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
constructor TGArgumentGetter.Create(const AArguments, ASeparator: String;
  AIgnoreCase: Boolean);
begin
  inherited Create();

  FArguments := AArguments;
  FSeparator := ASeparator;
  FIgnoreCase := AIgnoreCase;
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function TGArgumentGetter.GetBoolean(const Id: String): Boolean;
begin
	Result := StrToBool_(GetString(Id))
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function TGArgumentGetter.GetDateTime(const Id: String): TDateTime;
begin
	Result := StrToDateTimeDef(GetString(Id), iDate1900)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function TGArgumentGetter.GetDouble(const Id: String): Double;
begin
	Result := Valeur(GetString(Id))
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function TGArgumentGetter.GetInteger(const Id: String): Integer;
begin
  Result := ValeurI(GetString(Id))
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function TGArgumentGetter.GetObject(const Id: String): TObject;
begin
	if (not FIgnoreCase and (Pos(Id, FArguments) > 0))
  or (FIgnoreCase and (Pos(UpperCase(Id), UpperCase(FArguments)) > 0)) then
		Result := TObject(GetInteger(Id))
  else
   	Result := nil
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function TGArgumentGetter.GetString(const Id: String): String;
begin
	Result := VarToStr(GetValue(Id))
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
function TGArgumentGetter.GetValue(const Id: String): Variant;
var
	ArgTemp, Arg1, Arg2: String;
  posEqual: Integer;
begin
	Result := '';
  ArgTemp := FArguments;
  while (ArgTemp <> '') and (Result = '') do
  begin
    Arg1 := ReadTokenPipe(ArgTemp, Separator);
    if IgnoreCase then
     	Arg2 := UpperCase(Arg1)
    else
      Arg2 := Arg1;

    posEqual := Pos('=', Arg2);
   	if (Pos(Id, Arg2) > 0) and (posEqual <> 0) and (Trim(TGStrUtils.CopyLeft(Arg2, posEqual - 1)) = Id) then
   	  Result := Trim(Copy(Arg1, posEqual + 1, Length(Arg2)));
	end;
end;






{ ******************************************************************************
  TGArgUtils
  ****************************************************************************** }

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGArgUtils.GetBoolean(const Arguments, ArgumentID,
  Separator: String; IgnoreCase: Boolean): Boolean;
begin
  Result := ArgumentReader(Arguments, Separator, IgnoreCase).GetBoolean(ArgumentID)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGArgUtils.GetDateTime(const Arguments, ArgumentID,
  Separator: String; IgnoreCase: Boolean): tDateTime;
begin
  Result := ArgumentReader(Arguments, Separator, IgnoreCase).GetDateTime(ArgumentID)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGArgUtils.GetDouble(const Arguments, ArgumentID,
  Separator: String; IgnoreCase: Boolean): Double;
begin
  Result := ArgumentReader(Arguments, Separator, IgnoreCase).GetDouble(ArgumentID)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGArgUtils.GetInteger(const Arguments, ArgumentID,
  Separator: String; IgnoreCase: Boolean): Integer;
begin
  Result := ArgumentReader(Arguments, Separator, IgnoreCase).GetInteger(ArgumentID)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGArgUtils.GetString(const Arguments, ArgumentID,
  Separator: String; IgnoreCase: Boolean): String;
begin
  Result := ArgumentReader(Arguments, Separator, IgnoreCase).GetString(ArgumentID)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGArgUtils.GetTob(const Arguments, ArgumentID,
  Separator: String; IgnoreCase: Boolean): Tob;
begin
  Result := Tob(ArgumentReader(Arguments, Separator, IgnoreCase).GetObject(ArgumentID))
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGArgUtils.GetValue(const Arguments, ArgumentID,
  Separator: String; IgnoreCase: Boolean): String;
begin
  Result := ArgumentReader(Arguments, Separator, IgnoreCase).GetValue(ArgumentID)
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGArgUtils.MakeFromTob(const ArgumentID: String;
  const ATob: Tob; const Separator: String): String;
begin
  if (ArgumentID <> '') and Assigned(ATob) then
    Result := ArgumentID + Separator + IntToStr(LongInt(ATob))
  else
    Result := ''
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGArgUtils.Exists(const Arguments, ArgumentID, Separator: String): Boolean;
begin
  Result := IndexOfID(Arguments, ArgumentID, Separator) > -1
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGArgUtils.GetID(const Argument: String): String;
begin
  Result := UpperCase(Trim(TGStrUtils.CopyLeft(Argument, Pos('=', Argument) - 1)))
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
class function TGArgUtils.IndexOfID(const Arguments, ArgumentID, Separator: String): Integer;
var
  L: TStringList;
begin
  L := TStringList.Create();
  try
    if ArgumentID = '' then
      Result := -1
    else
    begin
      L.Text := StringReplace(Arguments, Separator, CrLf, [rfReplaceAll]);
      Result := L.IndexOfName(ArgumentID); //MyArg=MyVal;etc;
      if Result = -1 then
        Result := L.IndexOf(ArgumentID);   //MyArg;etc;
    end
  finally
    FreeAndNil(L);
  end;
end;






{ ******************************************************************************
  ExceptionG
  ****************************************************************************** }

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
procedure ExceptionG.SaveStatus(const Msg: WideString; NotifyIOError: Boolean;
  const IOError: TIOErr);
begin
  if V_Pgi.NbTransaction > 0 then
  begin
    V_Pgi.TransacErrorMessage := Msg;
    if NotifyIOError and (V_Pgi.ioError = oeOk) then
      V_Pgi.ioError := IOError
  end;

  if V_Pgi.FirstHShowMessage = '' then
    V_Pgi.FirstHShowMessage := Msg;
  V_Pgi.LastHShowMessage := Msg;

  if Trace.IsTraceErrorEnabled then
    Trace.TraceError(Self.ClassName, Self.ClassName + CrLf_ + Msg);
end;

{ ------------------------------------------------------------------------------
  ------------------------------------------------------------------------------ }
constructor ExceptionG.Create(const cbpIDError, SourceLocation: WideString;
  Args: array of Const; NotifyIOError: Boolean; const IOError: TIOErr);
var
  Msg: WideString;
begin
  if Length(Args) > 0 then
    Msg := Format_(cbpIDError, Args)
  else
    Msg := cbpIDError;
  Msg := StringReplace_(Msg, ';', CrLf_, [rfReplaceAll]);
  Msg := StringReplace_(Msg, '""', '@-x-@', [rfReplaceAll]);
  Msg := StringReplace_(Msg, '"', '', [rfReplaceAll]);
  Msg := StringReplace_(Msg, '@-x-@', '"', [rfReplaceAll]);
  Msg := SourceLocation + CrLf_ + Msg;
  SaveStatus(Msg, NotifyIOError, IOError);

  inherited Create(cbpIDError, SourceLocation, Args);
end;

end.
