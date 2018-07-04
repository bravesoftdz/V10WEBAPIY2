unit ZTypes;

interface

uses SaisComm ;

type RJal = record
            Code         : string[3] ;
            Nature       : string[3] ;
            ValideN      : string ;
            ValideN1     : string ;
            CInterdit    : string ;
            CAutomat     : string ;
            CContreP     : string ;
            ModeSaisie   : string ;
            bDevise      : Boolean ;
            LastNum      : Integer ;
            TotPerDebit  : Double ;     { Total période }
            TotPerCredit : Double ;
            TotFolDebit  : Double ;     { Total Folio en cours de saisie}
            TotFolCredit : Double ;
            TotDebDebit  : Double ;     { Total Folio avant la saisie }
            TotDebCredit : Double ;
            TotSaiDebit  : Double ;     { Total réel affiché en saisie }
            TotSaiCredit : Double ;
            TotVirDebit  : Double ;     { Total Equilibre virtuel }
            TotVirCredit : Double ;
            TotSavDebit  : Double ;     { Restore : Total dans la base }
            TotSavCredit : Double ;
            LastDateSais : TDateTime ;
            NbAuto       : Integer ;
            CurAuto      : Integer ;
            IncRef       : boolean ; // incremente automatique des references
            IncNum       : boolean ; // incremente le numero de folio
            EqAuto       : boolean ; // solde auto du folio
            NatCompl     : boolean ; // nature abrege ou complete
            NatDefaut    : string ;  // nature par defaut du journal
            Ferme        : boolean ;
            Accelerateur : boolean ;
            Lettrage     : boolean ;
            end ;

type ROpt = record
            Exo       : string[3] ;
            TypeExo   : TTypeExo ;
            MaxJour   : Word ;
            Month     : Word ;
            Year      : Word ;
            NewObj    : Boolean ;
            TypeBal   : string[3] ;
            PlanBal   : string[3] ;
            DebJour   : Word ;
            DebMonth  : Word ;
            DebYear   : Word ;
            ExoBal    : string[3] ;
            ExoAno    : string[3] ;
            DecalAno  : Integer ;
            WithAux   : Boolean ;
            WithRes   : Boolean ;
            bNoBilan  : Boolean ;
            bVisu     : Boolean ;
            bCalcRes  : Boolean ;
            bEcr      : Boolean ;
            NumPiece  : Integer ;
            bLettrage : Boolean ;
            QualifP   : Char ;
            end ;

type RPar = record
            bLibre       : Boolean ;
            bDevise      : Boolean ;
            DevLibre     : string ;
            bSoldeProg   : Boolean ;
            bRealTime    : Boolean ;
            end ;

type RDel = class
            NumLigne         : Integer ;
            NumEcheVent      : Integer ;
            DateModification : TDateTime ;
            Ax               : string ;
            DateComptable    : TDateTime ;
            NumFolio         : string ;
            QualifPiece      : string; {JP 20/11/06}
            Journal          : string; {JP 20/11/06}
            end ;

implementation

end.
