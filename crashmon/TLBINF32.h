// Created by Microsoft (R) C/C++ Compiler Version 12.00.8168.0 (c7c8cfd2).
//
// C++ source equivalent of Win32 type library TLBINF32.DLL
//
// had to be edited to get it to compile pfft -dz

#pragma once
#pragma pack(push, 8)

#include <comdef.h>

namespace TLI {

//
// Forward references and typedefs
//

struct __declspec(uuid("8b21774b-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ VarTypeInfo;
struct __declspec(uuid("8b217749-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ ParameterInfo;
struct __declspec(uuid("8b21774a-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ Parameters;
struct __declspec(uuid("8b217747-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ MemberInfo;
struct __declspec(uuid("8b217748-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ Members;
struct __declspec(uuid("8b217741-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ InterfaceInfo;
struct __declspec(uuid("8b217750-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ _BaseTypeInfos;
struct __declspec(uuid("8b217742-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ Interfaces;
struct __declspec(uuid("8b217743-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ CoClassInfo;
struct __declspec(uuid("8b217744-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ CoClasses;
struct __declspec(uuid("8b21774d-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ ConstantInfo;
struct __declspec(uuid("8b21774c-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ Constants;
struct __declspec(uuid("8b21774f-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ DeclarationInfo;
struct __declspec(uuid("8b21774e-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ Declarations;
struct __declspec(uuid("8b217751-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ _SearchHelper;
struct /* coclass */ SearchHelper;
struct __declspec(uuid("8b217759-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ TypeInfo;
struct __declspec(uuid("8b217745-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ _TypeLibInfo;
struct /* coclass */ TypeLibInfo;
struct __declspec(uuid("8b217756-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ SearchItem;
struct __declspec(uuid("8b217757-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ SearchResults;
struct __declspec(uuid("8b217758-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ ListBoxNotification;
struct __declspec(uuid("8b21775f-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ CustomSort;
struct __declspec(uuid("8b217760-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ CustomFilter;
struct __declspec(uuid("8b21775d-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ _TLIApplication;
struct /* coclass */ TLIApplication;
struct __declspec(uuid("8b21775a-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ TypeInfos;
struct __declspec(uuid("8b21775b-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ RecordInfo;
struct __declspec(uuid("8b21775c-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ Records;
struct __declspec(uuid("8b217761-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ IntrinsicAliasInfo;
struct __declspec(uuid("8b217762-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ IntrinsicAliases;
struct __declspec(uuid("8b217763-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ CustomData;
struct __declspec(uuid("8b217764-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ CustomDataCollection;
struct __declspec(uuid("8b217765-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ UnionInfo;
struct __declspec(uuid("8b217766-717d-11ce-ab5b-d41203c10000"))
/* dual interface */ Unions;

//
// Smart pointer typedef declarations
//

_COM_SMARTPTR_TYPEDEF(VarTypeInfo, __uuidof(VarTypeInfo));
_COM_SMARTPTR_TYPEDEF(ParameterInfo, __uuidof(ParameterInfo));
_COM_SMARTPTR_TYPEDEF(Parameters, __uuidof(Parameters));
_COM_SMARTPTR_TYPEDEF(MemberInfo, __uuidof(MemberInfo));
_COM_SMARTPTR_TYPEDEF(Members, __uuidof(Members));
_COM_SMARTPTR_TYPEDEF(InterfaceInfo, __uuidof(InterfaceInfo));
_COM_SMARTPTR_TYPEDEF(_BaseTypeInfos, __uuidof(_BaseTypeInfos));
_COM_SMARTPTR_TYPEDEF(Interfaces, __uuidof(Interfaces));
_COM_SMARTPTR_TYPEDEF(CoClassInfo, __uuidof(CoClassInfo));
_COM_SMARTPTR_TYPEDEF(CoClasses, __uuidof(CoClasses));
_COM_SMARTPTR_TYPEDEF(ConstantInfo, __uuidof(ConstantInfo));
_COM_SMARTPTR_TYPEDEF(Constants, __uuidof(Constants));
_COM_SMARTPTR_TYPEDEF(DeclarationInfo, __uuidof(DeclarationInfo));
_COM_SMARTPTR_TYPEDEF(Declarations, __uuidof(Declarations));
_COM_SMARTPTR_TYPEDEF(_SearchHelper, __uuidof(_SearchHelper));
_COM_SMARTPTR_TYPEDEF(TypeInfo, __uuidof(TypeInfo));
_COM_SMARTPTR_TYPEDEF(_TypeLibInfo, __uuidof(_TypeLibInfo));
_COM_SMARTPTR_TYPEDEF(SearchItem, __uuidof(SearchItem));
_COM_SMARTPTR_TYPEDEF(SearchResults, __uuidof(SearchResults));
_COM_SMARTPTR_TYPEDEF(ListBoxNotification, __uuidof(ListBoxNotification));
_COM_SMARTPTR_TYPEDEF(CustomSort, __uuidof(CustomSort));
_COM_SMARTPTR_TYPEDEF(CustomFilter, __uuidof(CustomFilter));
_COM_SMARTPTR_TYPEDEF(_TLIApplication, __uuidof(_TLIApplication));
_COM_SMARTPTR_TYPEDEF(TypeInfos, __uuidof(TypeInfos));
_COM_SMARTPTR_TYPEDEF(RecordInfo, __uuidof(RecordInfo));
_COM_SMARTPTR_TYPEDEF(Records, __uuidof(Records));
_COM_SMARTPTR_TYPEDEF(IntrinsicAliasInfo, __uuidof(IntrinsicAliasInfo));
_COM_SMARTPTR_TYPEDEF(IntrinsicAliases, __uuidof(IntrinsicAliases));
_COM_SMARTPTR_TYPEDEF(CustomData, __uuidof(CustomData));
_COM_SMARTPTR_TYPEDEF(CustomDataCollection, __uuidof(CustomDataCollection));
_COM_SMARTPTR_TYPEDEF(UnionInfo, __uuidof(UnionInfo));
_COM_SMARTPTR_TYPEDEF(Unions, __uuidof(Unions));

//
// Type library items
//

/*
enum TliErrors
{
    tliErrNoCurrentTypelib = -2147220991,
    tliErrCantLoadLibrary = -2147220990,
    tliErrTypeLibNotRegistered = -2147220989,
    tliErrSearchResultsChanged = -2147220988,
    tliErrNotApplicable = -2147220987,
    tliErrIncompatibleData = -2147220986,
    tliErrIncompatibleSearchType = -2147220985,
    tliErrIncompatibleTypeKind = -2147220984,
    tliErrInaccessibleImportLib = -2147220983,
    tliErrNoDefaultValue = -2147220982,
    tliErrDataNotAvailable = -2147220981,
    tliErrNotAnEntryPoint = -2147220980,
    tliErrStopFiltering = -2147220979,
    tliErrArrayBoundsNotAvailable = -2147220978,
    tliErrSearchResultsNotSorted = -2147220977,
    tliErrTypeNotArray = -2147220976
};

enum TypeFlags
{
    TYPEFLAG_NONE = 0,
    TYPEFLAG_FAPPOBJECT = 1,
    TYPEFLAG_FCANCREATE = 2,
    TYPEFLAG_FLICENSED = 4,
    TYPEFLAG_FPREDECLID = 8,
    TYPEFLAG_FHIDDEN = 16,
    TYPEFLAG_FCONTROL = 32,
    TYPEFLAG_FDUAL = 64,
    TYPEFLAG_FNONEXTENSIBLE = 128,
    TYPEFLAG_FOLEAUTOMATION = 256,
    TYPEFLAG_FRESTRICTED = 512,
    TYPEFLAG_FAGGREGATABLE = 1024,
    TYPEFLAG_FREPLACEABLE = 2048,
    TYPEFLAG_FDISPATCHABLE = 4096,
    TYPEFLAG_FREVERSEBIND = 8192,
    TYPEFLAG_FPROXY = 16384,
    TYPEFLAG_DEFAULTFILTER = 528,
    TYPEFLAG_COCLASSATTRIBUTES = 1599,
    TYPEFLAG_INTERFACEATTRIBUTES = 31680,
    TYPEFLAG_DISPATCHATTRIBUTES = 23232,
    TYPEFLAG_ALIASATTRIBUTES = 528,
    TYPEFLAG_MODULEATTRIBUTES = 528,
    TYPEFLAG_ENUMATTRIBUTES = 528,
    TYPEFLAG_RECORDATTRIBUTES = 528,
    TYPEFLAG_UNIONATTRIBUTES = 528
};

enum ImplTypeFlags
{
    IMPLTYPEFLAG_FDEFAULT = 1,
    IMPLTYPEFLAG_FSOURCE = 2,
    IMPLTYPEFLAG_FRESTRICTED = 4,
    IMPLTYPEFLAG_FDEFAULTVTABLE = 8
};



enum IDLFlags
{
    IDLFLAG_NONE = 0,
    IDLFLAG_FIN = 1,
    IDLFLAG_FOUT = 2,
    IDLFLAG_FLCID = 4,
    IDLFLAG_FRETVAL = 8
};

enum ParamFlags
{
    PARAMFLAG_NONE = 0,
    PARAMFLAG_FIN = 1,
    PARAMFLAG_FOUT = 2,
    PARAMFLAG_FLCID = 4,
    PARAMFLAG_FRETVAL = 8,
    PARAMFLAG_FOPT = 16,
    PARAMFLAG_FHASDEFAULT = 32,
    PARAMFLAG_FHASCUSTDATA = 64
};



enum TliSearchTypes
{
    tliStDefault = 4096,
    tliStClasses = 1,
    tliStEvents = 2,
    tliStConstants = 4,
    tliStDeclarations = 8,
    tliStAppObject = 16,
    tliStRecords = 32,
    tliStIntrinsicAliases = 64,
    tliStUnions = 128,
    tliStAll = 239
};

enum TliWindowTypes
{
    tliWtListBox = 0,
    tliWtComboBox = 1
};

enum TliItemDataTypes
{
    tliIdtMemberID = 0,
    tliIdtInvokeKinds = 1
};
*/

struct __declspec(uuid("8b21774b-717d-11ce-ab5b-d41203c10000"))
VarTypeInfo : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct VarTypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall get__OldVarType ( ) = 0;
    virtual HRESULT __stdcall get_TypeInfo (
        struct TypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_TypeInfoNumber (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_TypedVariant (
        VARIANT * retVal ) = 0;
    virtual HRESULT __stdcall get_IsExternalType (
        VARIANT_BOOL * retVal ) = 0;
    virtual HRESULT __stdcall get_TypeLibInfoExternal (
        struct _TypeLibInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_PointerLevel (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_VarType (
        enum TliVarType * retVal ) = 0;
    virtual HRESULT __stdcall ArrayBounds (
        SAFEARRAY * * Bounds,
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_ElementPointerLevel (
        short * retVal ) = 0;
};

struct __declspec(uuid("8b217749-717d-11ce-ab5b-d41203c10000"))
ParameterInfo : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct ParameterInfo * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall get_Name (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_Optional (
        VARIANT_BOOL * retVal ) = 0;
    virtual HRESULT __stdcall get__OldFlags ( ) = 0;
    virtual HRESULT __stdcall get_VarTypeInfo (
        struct VarTypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_Default (
        VARIANT_BOOL * retVal ) = 0;
    virtual HRESULT __stdcall get_DefaultValue (
        VARIANT * retVal ) = 0;
    virtual HRESULT __stdcall get_HasCustomData (
        VARIANT_BOOL * retVal ) = 0;
    virtual HRESULT __stdcall get_CustomDataCollection (
        struct CustomDataCollection * * retVal ) = 0;
    virtual HRESULT __stdcall get_Flags (
        enum ParamFlags * retVal ) = 0;
};

struct __declspec(uuid("8b21774a-717d-11ce-ab5b-d41203c10000"))
Parameters : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct Parameters * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall _NewEnum (
        IUnknown * * retVal ) = 0;
    virtual HRESULT __stdcall get_Item (
        short Index,
        struct ParameterInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_Count (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_OptionalCount (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_DefaultCount (
        short * retVal ) = 0;
};

struct __declspec(uuid("8b217747-717d-11ce-ab5b-d41203c10000"))
MemberInfo : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct MemberInfo * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall get_Name (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get__OldHelpString ( ) = 0;
    virtual HRESULT __stdcall get_HelpContext (
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpFile (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_Parameters (
        struct Parameters * * retVal ) = 0;
    virtual HRESULT __stdcall get_ReturnType (
        struct VarTypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeMask (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeStrings (
        SAFEARRAY * * AttributeArray,
        short * retVal ) = 0;
    virtual HRESULT __stdcall get__OldDescKind ( ) = 0;
    virtual HRESULT __stdcall get_Value (
        VARIANT * retVal ) = 0;
    virtual HRESULT __stdcall get_MemberId (
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_VTableOffset (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_InvokeKind (
        enum InvokeKinds * retVal ) = 0;
    virtual HRESULT __stdcall get_CustomDataCollection (
        struct CustomDataCollection * * retVal ) = 0;
    virtual HRESULT __stdcall get_DescKind (
        enum DescKinds * retVal ) = 0;
    virtual HRESULT __stdcall GetDllEntry (
        BSTR * DllName,
        BSTR * EntryName,
        short * Ordinal ) = 0;
    virtual HRESULT __stdcall get_HelpString (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringDll (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringContext (
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_CallConv (
        enum CallConvs * retVal ) = 0;
};

struct __declspec(uuid("8b217748-717d-11ce-ab5b-d41203c10000"))
Members : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct Members * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall _NewEnum (
        IUnknown * * retVal ) = 0;
    virtual HRESULT __stdcall get_Item (
        short Index,
        struct MemberInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_Count (
        short * retVal ) = 0;
    virtual HRESULT __stdcall put_FuncFilter (
        enum FuncFlags retVal ) = 0;
    virtual HRESULT __stdcall get_FuncFilter (
        enum FuncFlags * retVal ) = 0;
    virtual HRESULT __stdcall put_VarFilter (
        enum VarFlags retVal ) = 0;
    virtual HRESULT __stdcall get_VarFilter (
        enum VarFlags * retVal ) = 0;
    virtual HRESULT __stdcall _OldFillList ( ) = 0;
    virtual HRESULT __stdcall get_GetFilteredMembers (
        VARIANT_BOOL ShowUnderscore,
        struct SearchResults * * retVal ) = 0;
    virtual HRESULT __stdcall GetFilteredMembersDirect (
        int hWnd,
        enum TliWindowTypes WindowType,
        enum TliItemDataTypes ItemDataType,
        VARIANT_BOOL ShowUnderscore,
        short * retVal ) = 0;
};

struct __declspec(uuid("8b217741-717d-11ce-ab5b-d41203c10000"))
InterfaceInfo : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct InterfaceInfo * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall get_Name (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_GUID (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get__OldHelpString ( ) = 0;
    virtual HRESULT __stdcall get_HelpContext (
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpFile (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeMask (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeStrings (
        SAFEARRAY * * AttributeArray,
        short * retVal ) = 0;
    virtual HRESULT __stdcall get__OldTypeKind ( ) = 0;
    virtual HRESULT __stdcall get_TypeKindString (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_TypeInfoNumber (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_VTableInterface (
        struct InterfaceInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_GetMember (
        VARIANT Index,
        struct MemberInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_Members (
        struct Members * * retVal ) = 0;
    virtual HRESULT __stdcall get_Parent (
        struct _TypeLibInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_ImpliedInterfaces (
        struct Interfaces * * retVal ) = 0;
    virtual HRESULT __stdcall _DefaultInterface ( ) = 0;
    virtual HRESULT __stdcall _DefaultEventInterface ( ) = 0;
    virtual HRESULT __stdcall get_TypeKind (
        enum TypeKinds * retVal ) = 0;
    virtual HRESULT __stdcall get_ResolvedType (
        struct VarTypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_CustomDataCollection (
        struct CustomDataCollection * * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpString (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_ITypeInfo (
        IUnknown * * retVal ) = 0;
    virtual HRESULT __stdcall get_MajorVersion (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_MinorVersion (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringDll (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringContext (
        long * retVal ) = 0;
};

struct __declspec(uuid("8b217750-717d-11ce-ab5b-d41203c10000"))
_BaseTypeInfos : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct TypeInfos * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall _NewEnum (
        IUnknown * * retVal ) = 0;
    virtual HRESULT __stdcall get_Count (
        short * retVal ) = 0;
};

struct __declspec(uuid("8b217742-717d-11ce-ab5b-d41203c10000"))
Interfaces : _BaseTypeInfos
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall get_Item (
        short Index,
        struct InterfaceInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_IndexedItem (
        short TypeInfoNumber,
        struct InterfaceInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_NamedItem (
        BSTR * TypeInfoName,
        struct InterfaceInfo * * retVal ) = 0;
};

struct __declspec(uuid("8b217743-717d-11ce-ab5b-d41203c10000"))
CoClassInfo : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct CoClassInfo * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall get_Name (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_GUID (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get__OldHelpString ( ) = 0;
    virtual HRESULT __stdcall get_HelpContext (
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpFile (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeMask (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeStrings (
        SAFEARRAY * * AttributeArray,
        short * retVal ) = 0;
    virtual HRESULT __stdcall get__OldTypeKind ( ) = 0;
    virtual HRESULT __stdcall get_TypeKindString (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_TypeInfoNumber (
        short * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_VTableInterface ( ) = 0;
    virtual HRESULT __stdcall _placeholder_GetMember ( ) = 0;
    virtual HRESULT __stdcall _placeholder_Members ( ) = 0;
    virtual HRESULT __stdcall get_Parent (
        struct _TypeLibInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_Interfaces (
        struct Interfaces * * retVal ) = 0;
    virtual HRESULT __stdcall get_DefaultInterface (
        struct InterfaceInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_DefaultEventInterface (
        struct InterfaceInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_TypeKind (
        enum TypeKinds * retVal ) = 0;
    virtual HRESULT __stdcall get_ResolvedType (
        struct VarTypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_CustomDataCollection (
        struct CustomDataCollection * * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpString (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_ITypeInfo (
        IUnknown * * retVal ) = 0;
    virtual HRESULT __stdcall get_MajorVersion (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_MinorVersion (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringDll (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringContext (
        long * retVal ) = 0;
};

struct __declspec(uuid("8b217744-717d-11ce-ab5b-d41203c10000"))
CoClasses : _BaseTypeInfos
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall get_Item (
        short Index,
        struct CoClassInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_IndexedItem (
        short TypeInfoNumber,
        struct CoClassInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_NamedItem (
        BSTR * TypeInfoName,
        struct CoClassInfo * * retVal ) = 0;
};

struct __declspec(uuid("8b21774d-717d-11ce-ab5b-d41203c10000"))
ConstantInfo : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct TypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall get_Name (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_GUID (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get__OldHelpString ( ) = 0;
    virtual HRESULT __stdcall get_HelpContext (
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpFile (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeMask (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeStrings (
        SAFEARRAY * * AttributeArray,
        short * retVal ) = 0;
    virtual HRESULT __stdcall get__OldTypeKind ( ) = 0;
    virtual HRESULT __stdcall get_TypeKindString (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_TypeInfoNumber (
        short * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_VTableInterface ( ) = 0;
    virtual HRESULT __stdcall get_GetMember (
        VARIANT Index,
        struct MemberInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_Members (
        struct Members * * retVal ) = 0;
    virtual HRESULT __stdcall get_Parent (
        struct _TypeLibInfo * * retVal ) = 0;
    virtual HRESULT __stdcall _ImpliedInterfaces ( ) = 0;
    virtual HRESULT __stdcall _DefaultInterface ( ) = 0;
    virtual HRESULT __stdcall _DefaultEventInterface ( ) = 0;
    virtual HRESULT __stdcall get_TypeKind (
        enum TypeKinds * retVal ) = 0;
    virtual HRESULT __stdcall get_ResolvedType (
        struct VarTypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_CustomDataCollection (
        struct CustomDataCollection * * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpString (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_ITypeInfo (
        IUnknown * * retVal ) = 0;
    virtual HRESULT __stdcall get_MajorVersion (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_MinorVersion (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringDll (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringContext (
        long * retVal ) = 0;
};

struct __declspec(uuid("8b21774c-717d-11ce-ab5b-d41203c10000"))
Constants : _BaseTypeInfos
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall get_Item (
        short Index,
        struct ConstantInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_IndexedItem (
        short TypeInfoNumber,
        struct ConstantInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_NamedItem (
        BSTR * TypeInfoName,
        struct ConstantInfo * * retVal ) = 0;
};

struct __declspec(uuid("8b21774f-717d-11ce-ab5b-d41203c10000"))
DeclarationInfo : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct DeclarationInfo * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall get_Name (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_GUID (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get__OldHelpString ( ) = 0;
    virtual HRESULT __stdcall get_HelpContext (
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpFile (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeMask (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeStrings (
        SAFEARRAY * * AttributeArray,
        short * retVal ) = 0;
    virtual HRESULT __stdcall get__OldTypeKind ( ) = 0;
    virtual HRESULT __stdcall get_TypeKindString (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_TypeInfoNumber (
        short * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_VTableInterface ( ) = 0;
    virtual HRESULT __stdcall get_GetMember (
        VARIANT Index,
        struct MemberInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_Members (
        struct Members * * retVal ) = 0;
    virtual HRESULT __stdcall get_Parent (
        struct _TypeLibInfo * * retVal ) = 0;
    virtual HRESULT __stdcall _ImpliedInterfaces ( ) = 0;
    virtual HRESULT __stdcall _DefaultInterface ( ) = 0;
    virtual HRESULT __stdcall _DefaultEventInterface ( ) = 0;
    virtual HRESULT __stdcall get_TypeKind (
        enum TypeKinds * retVal ) = 0;
    virtual HRESULT __stdcall get_ResolvedType (
        struct VarTypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_CustomDataCollection (
        struct CustomDataCollection * * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpString (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_ITypeInfo (
        IUnknown * * retVal ) = 0;
    virtual HRESULT __stdcall get_MajorVersion (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_MinorVersion (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringDll (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringContext (
        long * retVal ) = 0;
};

struct __declspec(uuid("8b21774e-717d-11ce-ab5b-d41203c10000"))
Declarations : _BaseTypeInfos
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall get_Item (
        short Index,
        struct DeclarationInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_IndexedItem (
        short TypeInfoNumber,
        struct DeclarationInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_NamedItem (
        BSTR * TypeInfoName,
        struct DeclarationInfo * * retVal ) = 0;
};

struct __declspec(uuid("8b217751-717d-11ce-ab5b-d41203c10000"))
_SearchHelper : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct _TypeLibInfo * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall _OldInit ( ) = 0;
    virtual HRESULT __stdcall get_CheckHaveMatch (
        BSTR Name,
        VARIANT_BOOL * retVal ) = 0;
    virtual HRESULT __stdcall get_Init (
        enum SysKinds SysKind,
        long LCID,
        short GrowSize ) = 0;
};

struct __declspec(uuid("8b217752-717d-11ce-ab5b-d41203c10000"))
SearchHelper;
    // [ default ] interface _SearchHelper

struct __declspec(uuid("8b217759-717d-11ce-ab5b-d41203c10000"))
TypeInfo : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct TypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall get_Name (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_GUID (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get__OldHelpString ( ) = 0;
    virtual HRESULT __stdcall get_HelpContext (
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpFile (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeMask (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeStrings (
        SAFEARRAY * * AttributeArray,
        short * retVal ) = 0;
    virtual HRESULT __stdcall get__OldTypeKind ( ) = 0;
    virtual HRESULT __stdcall get_TypeKindString (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_TypeInfoNumber (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_VTableInterface (
        struct InterfaceInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_GetMember (
        VARIANT Index,
        struct MemberInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_Members (
        struct Members * * retVal ) = 0;
    virtual HRESULT __stdcall get_Parent (
        struct _TypeLibInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_Interfaces (
        struct Interfaces * * retVal ) = 0;
    virtual HRESULT __stdcall get_DefaultInterface (
        struct InterfaceInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_DefaultEventInterface (
        struct InterfaceInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_TypeKind (
        enum TypeKinds * retVal ) = 0;
    virtual HRESULT __stdcall get_ResolvedType (
        struct VarTypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_CustomDataCollection (
        struct CustomDataCollection * * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpString (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_ITypeInfo (
        IUnknown * * retVal ) = 0;
    virtual HRESULT __stdcall get_MajorVersion (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_MinorVersion (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringDll (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringContext (
        long * retVal ) = 0;
};

struct __declspec(uuid("8b217745-717d-11ce-ab5b-d41203c10000"))
_TypeLibInfo : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct _TypeLibInfo * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall get_ContainingFile (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall put_ContainingFile (
        BSTR retVal ) = 0;
    virtual HRESULT __stdcall LoadRegTypeLib (
        BSTR TypeLibGuid,
        short MajorVersion,
        short MinorVersion,
        long LCID ) = 0;
    virtual HRESULT __stdcall get_Name (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get__OldHelpString ( ) = 0;
    virtual HRESULT __stdcall get_HelpContext (
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpFile (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_GUID (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_LCID (
        long * retVal ) = 0;
    virtual HRESULT __stdcall get__OldSysKind ( ) = 0;
    virtual HRESULT __stdcall get_MajorVersion (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_MinorVersion (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeMask (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeStrings (
        SAFEARRAY * * AttributeArray,
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_CoClasses (
        struct CoClasses * * retVal ) = 0;
    virtual HRESULT __stdcall get_Interfaces (
        struct Interfaces * * retVal ) = 0;
    virtual HRESULT __stdcall get_Constants (
        struct Constants * * retVal ) = 0;
    virtual HRESULT __stdcall get_Declarations (
        struct Declarations * * retVal ) = 0;
    virtual HRESULT __stdcall get_TypeInfoCount (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get__OldGetTypeKind ( ) = 0;
    virtual HRESULT __stdcall get_GetTypeInfo (
        VARIANT * Index,
        struct TypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_GetTypeInfoNumber (
        BSTR Name,
        short * retVal ) = 0;
    virtual HRESULT __stdcall IsSameLibrary (
        struct _TypeLibInfo * CheckLib,
        VARIANT_BOOL * retVal ) = 0;
    virtual HRESULT __stdcall _OldResetSearchCriteria ( ) = 0;
    virtual HRESULT __stdcall _OldGetTypesWithMember ( ) = 0;
    virtual HRESULT __stdcall _OldGetMembersWithSubString ( ) = 0;
    virtual HRESULT __stdcall _OldGetTypesWithSubString ( ) = 0;
    virtual HRESULT __stdcall _OldCaseTypeName ( ) = 0;
    virtual HRESULT __stdcall _OldCaseMemberName ( ) = 0;
    virtual HRESULT __stdcall _OldFillTypesList ( ) = 0;
    virtual HRESULT __stdcall _OldFillTypesCombo ( ) = 0;
    virtual HRESULT __stdcall _OldFillMemberList ( ) = 0;
    virtual HRESULT __stdcall _OldAddClassTypeToList ( ) = 0;
    virtual HRESULT __stdcall put_AppObjString (
        BSTR retVal ) = 0;
    virtual HRESULT __stdcall put_LibNum (
        short retVal ) = 0;
    virtual HRESULT __stdcall get_ShowLibName (
        VARIANT_BOOL * retVal ) = 0;
    virtual HRESULT __stdcall put_ShowLibName (
        VARIANT_BOOL retVal ) = 0;
    virtual HRESULT __stdcall putref__OldListBoxNotification (
        struct ListBoxNotification * _arg1 ) = 0;
    virtual HRESULT __stdcall get_GetTypeKind (
        short TypeInfoNumber,
        enum TypeKinds * retVal ) = 0;
    virtual HRESULT __stdcall get_SysKind (
        enum SysKinds * retVal ) = 0;
    virtual HRESULT __stdcall get_SearchDefault (
        enum TliSearchTypes * retVal ) = 0;
    virtual HRESULT __stdcall put_SearchDefault (
        enum TliSearchTypes retVal ) = 0;
    virtual HRESULT __stdcall CaseTypeName (
        BSTR * bstrName,
        enum TliSearchTypes SearchType,
        enum TliSearchTypes * retVal ) = 0;
    virtual HRESULT __stdcall CaseMemberName (
        BSTR * bstrName,
        enum TliSearchTypes SearchType,
        VARIANT_BOOL * retVal ) = 0;
    virtual HRESULT __stdcall ResetSearchCriteria (
        enum TypeFlags TypeFilter,
        VARIANT_BOOL IncludeEmptyTypes,
        VARIANT_BOOL ShowUnderscore ) = 0;
    virtual HRESULT __stdcall GetTypesWithMember (
        BSTR MemberName,
        struct SearchResults * * StartResults,
        enum TliSearchTypes SearchType,
        VARIANT_BOOL Sort,
        VARIANT_BOOL ShowUnderscore,
        struct SearchResults * * retVal ) = 0;
    virtual HRESULT __stdcall GetTypesWithMemberDirect (
        BSTR MemberName,
        int hWnd,
        enum TliWindowTypes WindowType,
        enum TliSearchTypes SearchType,
        VARIANT_BOOL ShowUnderscore,
        short * retVal ) = 0;
    virtual HRESULT __stdcall GetMembersWithSubString (
        BSTR SubString,
        struct SearchResults * * StartResults,
        enum TliSearchTypes SearchType,
        VARIANT_BOOL SearchMiddle,
        struct _SearchHelper * * Helper,
        VARIANT_BOOL Sort,
        VARIANT_BOOL ShowUnderscore,
        struct SearchResults * * retVal ) = 0;
    virtual HRESULT __stdcall GetMembersWithSubStringDirect (
        BSTR SubString,
        int hWnd,
        enum TliWindowTypes WindowType,
        enum TliSearchTypes SearchType,
        VARIANT_BOOL SearchMiddle,
        struct _SearchHelper * * Helper,
        VARIANT_BOOL ShowUnderscore,
        short * retVal ) = 0;
    virtual HRESULT __stdcall GetTypesWithSubString (
        BSTR SubString,
        struct SearchResults * * StartResults,
        enum TliSearchTypes SearchType,
        VARIANT_BOOL SearchMiddle,
        VARIANT_BOOL Sort,
        struct SearchResults * * retVal ) = 0;
    virtual HRESULT __stdcall GetTypesWithSubStringDirect (
        BSTR SubString,
        int hWnd,
        enum TliWindowTypes WindowType,
        enum TliSearchTypes SearchType,
        VARIANT_BOOL SearchMiddle,
        short * retVal ) = 0;
    virtual HRESULT __stdcall GetTypes (
        struct SearchResults * * StartResults,
        enum TliSearchTypes SearchType,
        VARIANT_BOOL Sort,
        struct SearchResults * * retVal ) = 0;
    virtual HRESULT __stdcall GetTypesDirect (
        int hWnd,
        enum TliWindowTypes WindowType,
        enum TliSearchTypes SearchType,
        short * retVal ) = 0;
    virtual HRESULT __stdcall GetMembers (
        long SearchData,
        VARIANT_BOOL ShowUnderscore,
        struct SearchResults * * retVal ) = 0;
    virtual HRESULT __stdcall GetMembersDirect (
        long SearchData,
        int hWnd,
        enum TliWindowTypes WindowType,
        enum TliItemDataTypes ItemDataType,
        VARIANT_BOOL ShowUnderscore,
        short * retVal ) = 0;
    virtual HRESULT __stdcall SetMemberFilters (
        enum FuncFlags FuncFilter,
        enum VarFlags VarFilter ) = 0;
    virtual HRESULT __stdcall MakeSearchData (
        BSTR TypeInfoName,
        enum TliSearchTypes SearchType,
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_TypeInfos (
        struct TypeInfos * * retVal ) = 0;
    virtual HRESULT __stdcall get_Records (
        struct Records * * retVal ) = 0;
    virtual HRESULT __stdcall get_IntrinsicAliases (
        struct IntrinsicAliases * * retVal ) = 0;
    virtual HRESULT __stdcall get_CustomDataCollection (
        struct CustomDataCollection * * retVal ) = 0;
    virtual HRESULT __stdcall GetMemberInfo (
        long SearchData,
        enum InvokeKinds InvokeKinds,
        long MemberId,
        BSTR MemberName,
        struct MemberInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_Unions (
        struct Unions * * retVal ) = 0;
    virtual HRESULT __stdcall AddTypes (
        SAFEARRAY * * TypeInfoNumbers,
        struct SearchResults * * StartResults,
        enum TliSearchTypes SearchType,
        VARIANT_BOOL Sort,
        struct SearchResults * * retVal ) = 0;
    virtual HRESULT __stdcall AddTypesDirect (
        SAFEARRAY * * TypeInfoNumbers,
        int hWnd,
        enum TliWindowTypes WindowType,
        enum TliSearchTypes SearchType,
        short * retVal ) = 0;
    virtual HRESULT __stdcall FreeSearchCriteria ( ) = 0;
    virtual HRESULT __stdcall Register (
        BSTR HelpDir ) = 0;
    virtual HRESULT __stdcall UnRegister ( ) = 0;
    virtual HRESULT __stdcall get_HelpString (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_AppObjString (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_LibNum (
        short * retVal ) = 0;
    virtual HRESULT __stdcall GetMembersWithSubStringEx (
        BSTR SubString,
        SAFEARRAY * * InvokeGroupings,
        struct SearchResults * * StartResults,
        enum TliSearchTypes SearchType,
        VARIANT_BOOL SearchMiddle,
        VARIANT_BOOL Sort,
        VARIANT_BOOL ShowUnderscore,
        struct SearchResults * * retVal ) = 0;
    virtual HRESULT __stdcall GetTypesWithMemberEx (
        BSTR MemberName,
        enum InvokeKinds InvokeKind,
        struct SearchResults * * StartResults,
        enum TliSearchTypes SearchType,
        VARIANT_BOOL Sort,
        VARIANT_BOOL ShowUnderscore,
        struct SearchResults * * retVal ) = 0;
    virtual HRESULT __stdcall get_ITypeLib (
        IUnknown * * retVal ) = 0;
    virtual HRESULT __stdcall putref_ITypeLib (
        IUnknown * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringDll (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringContext (
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_BestEquivalentType (
        BSTR TypeInfoName,
        BSTR * retVal ) = 0;
};

struct __declspec(uuid("8b217746-717d-11ce-ab5b-d41203c10000"))
TypeLibInfo;
    // [ default ] interface _TypeLibInfo

struct __declspec(uuid("8b217756-717d-11ce-ab5b-d41203c10000"))
SearchItem : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct SearchItem * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall get_Name (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_SearchData (
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_TypeInfoNumber (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get__OldLibNum (
        unsigned char * retVal ) = 0;
    virtual HRESULT __stdcall get_SearchType (
        enum TliSearchTypes * retVal ) = 0;
    virtual HRESULT __stdcall get_MemberId (
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_InvokeKinds (
        enum InvokeKinds * retVal ) = 0;
    virtual HRESULT __stdcall get_NamePtrW (
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_LibNum (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_Constant (
        VARIANT_BOOL * retVal ) = 0;
    virtual HRESULT __stdcall get_Hidden (
        VARIANT_BOOL * retVal ) = 0;
    virtual HRESULT __stdcall get_InvokeGroup (
        short * retVal ) = 0;
};

struct __declspec(uuid("8b217757-717d-11ce-ab5b-d41203c10000"))
SearchResults : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct SearchResults * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall _NewEnum (
        IUnknown * * retVal ) = 0;
    virtual HRESULT __stdcall get__OldItem (
        short Index,
        struct SearchItem * * retVal ) = 0;
    virtual HRESULT __stdcall get__OldCount (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_Sorted (
        VARIANT_BOOL * retVal ) = 0;
    virtual HRESULT __stdcall Sort (
        struct CustomSort * CustomSort ) = 0;
    virtual HRESULT __stdcall Filter (
        struct CustomFilter * CustomFilter,
        struct SearchResults * * AppendExtractedTo,
        struct SearchItem * StartAfter,
        struct SearchResults * * retVal ) = 0;
    virtual HRESULT __stdcall get_Item (
        long Index,
        struct SearchItem * * retVal ) = 0;
    virtual HRESULT __stdcall get_Count (
        long * retVal ) = 0;
    virtual HRESULT __stdcall LocateSorted (
        struct CustomSort * CustomSort,
        long * retVal ) = 0;
    virtual HRESULT __stdcall Locate (
        BSTR SearchString,
        struct CustomSort * CustomSort,
        long StartAfter,
        long * retVal ) = 0;
};

struct __declspec(uuid("8b217758-717d-11ce-ab5b-d41203c10000"))
ListBoxNotification : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall OnAddString (
        long lpstr,
        VARIANT_BOOL fUnicode ) = 0;
};

struct __declspec(uuid("8b21775f-717d-11ce-ab5b-d41203c10000"))
CustomSort : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Compare (
        struct SearchItem * Item1,
        struct SearchItem * Item2,
        long * Compare ) = 0;
};

enum TliCustomFilterAction
{
    tliCfaLeave = 0,
    tliCfaDuplicate = 1,
    tliCfaExtract = 2,
    tliCfaDelete = 3
};

struct __declspec(uuid("8b217760-717d-11ce-ab5b-d41203c10000"))
CustomFilter : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Visit (
        struct SearchItem * Item,
        enum TliCustomFilterAction * Action ) = 0;
};

struct __declspec(uuid("8b21775d-717d-11ce-ab5b-d41203c10000"))
_TLIApplication : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct _TLIApplication * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall TypeLibInfoFromFile (
        BSTR FileName,
        struct _TypeLibInfo * * retVal ) = 0;
    virtual HRESULT __stdcall InterfaceInfoFromObject (
        IDispatch * Object,
        struct InterfaceInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_ListBoxNotification (
        struct ListBoxNotification * * retVal ) = 0;
    virtual HRESULT __stdcall putref_ListBoxNotification (
        struct ListBoxNotification * retVal ) = 0;
    virtual HRESULT __stdcall get_ResolveAliases (
        VARIANT_BOOL * retVal ) = 0;
    virtual HRESULT __stdcall put_ResolveAliases (
        VARIANT_BOOL retVal ) = 0;
    virtual HRESULT __stdcall InvokeHook (
        IDispatch * Object,
        VARIANT NameOrID,
        enum InvokeKinds InvokeKind,
        SAFEARRAY * * ReverseArgList,
        VARIANT * retVal ) = 0;
    virtual HRESULT __stdcall InvokeHookArray (
        IDispatch * Object,
        VARIANT NameOrID,
        enum InvokeKinds InvokeKind,
        SAFEARRAY * * ReverseArgList,
        VARIANT * retVal ) = 0;
    virtual HRESULT __stdcall InvokeHookSub (
        IDispatch * Object,
        VARIANT NameOrID,
        enum InvokeKinds InvokeKind,
        SAFEARRAY * * ReverseArgList ) = 0;
    virtual HRESULT __stdcall InvokeHookArraySub (
        IDispatch * Object,
        VARIANT NameOrID,
        enum InvokeKinds InvokeKind,
        SAFEARRAY * * ReverseArgList ) = 0;
    virtual HRESULT __stdcall ClassInfoFromObject (
        IUnknown * Object,
        struct TypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall InvokeID (
        IDispatch * Object,
        BSTR Name,
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_InvokeLCID (
        long * retVal ) = 0;
    virtual HRESULT __stdcall put_InvokeLCID (
        long retVal ) = 0;
    virtual HRESULT __stdcall TypeInfoFromITypeInfo (
        IUnknown * ptinfo,
        struct TypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall TypeLibInfoFromITypeLib (
        IUnknown * pITypeLib,
        struct _TypeLibInfo * * retVal ) = 0;
    virtual HRESULT __stdcall TypeLibInfoFromRegistry (
        BSTR TypeLibGuid,
        short MajorVersion,
        short MinorVersion,
        long LCID,
        struct _TypeLibInfo * * retVal ) = 0;
    virtual HRESULT __stdcall TypeInfoFromRecordVariant (
        VARIANT * RecordVariant,
        struct TypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_RecordField (
        VARIANT * RecordVariant,
        BSTR * FieldName,
        VARIANT * retVal ) = 0;
    virtual HRESULT __stdcall put_RecordField (
        VARIANT * RecordVariant,
        BSTR * FieldName,
        VARIANT * retVal ) = 0;
    virtual HRESULT __stdcall putref_RecordField (
        VARIANT * RecordVariant,
        BSTR * FieldName,
        VARIANT * retVal ) = 0;
};

struct __declspec(uuid("8b21775e-717d-11ce-ab5b-d41203c10000"))
TLIApplication;
    // [ default ] interface _TLIApplication

struct __declspec(uuid("8b21775a-717d-11ce-ab5b-d41203c10000"))
TypeInfos : _BaseTypeInfos
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall get_Item (
        short Index,
        struct TypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_IndexedItem (
        short TypeInfoNumber,
        struct TypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_NamedItem (
        BSTR * TypeInfoName,
        struct TypeInfo * * retVal ) = 0;
};

struct __declspec(uuid("8b21775b-717d-11ce-ab5b-d41203c10000"))
RecordInfo : ConstantInfo
{};

struct __declspec(uuid("8b21775c-717d-11ce-ab5b-d41203c10000"))
Records : _BaseTypeInfos
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall get_Item (
        short Index,
        struct RecordInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_IndexedItem (
        short TypeInfoNumber,
        struct RecordInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_NamedItem (
        BSTR * TypeInfoName,
        struct RecordInfo * * retVal ) = 0;
};

struct __declspec(uuid("8b217761-717d-11ce-ab5b-d41203c10000"))
IntrinsicAliasInfo : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct IntrinsicAliasInfo * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall get_Name (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_GUID (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get__OldHelpString ( ) = 0;
    virtual HRESULT __stdcall get_HelpContext (
        long * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpFile (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeMask (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_AttributeStrings (
        SAFEARRAY * * AttributeArray,
        short * retVal ) = 0;
    virtual HRESULT __stdcall get__OldTypeKind ( ) = 0;
    virtual HRESULT __stdcall get_TypeKindString (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_TypeInfoNumber (
        short * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_VTableInterface ( ) = 0;
    virtual HRESULT __stdcall _placeholder_GetMember ( ) = 0;
    virtual HRESULT __stdcall _placeholder_Members ( ) = 0;
    virtual HRESULT __stdcall get_Parent (
        struct _TypeLibInfo * * retVal ) = 0;
    virtual HRESULT __stdcall _ImpliedInterfaces ( ) = 0;
    virtual HRESULT __stdcall _DefaultInterface ( ) = 0;
    virtual HRESULT __stdcall _DefaultEventInterface ( ) = 0;
    virtual HRESULT __stdcall get_TypeKind (
        enum TypeKinds * retVal ) = 0;
    virtual HRESULT __stdcall get_ResolvedType (
        struct VarTypeInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_CustomDataCollection (
        struct CustomDataCollection * * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpString (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_ITypeInfo (
        IUnknown * * retVal ) = 0;
    virtual HRESULT __stdcall get_MajorVersion (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_MinorVersion (
        short * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringDll (
        long LCID,
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_HelpStringContext (
        long * retVal ) = 0;
};

struct __declspec(uuid("8b217762-717d-11ce-ab5b-d41203c10000"))
IntrinsicAliases : _BaseTypeInfos
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall get_Item (
        short Index,
        struct IntrinsicAliasInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_IndexedItem (
        short TypeInfoNumber,
        struct IntrinsicAliasInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_NamedItem (
        BSTR * TypeInfoName,
        struct IntrinsicAliasInfo * * retVal ) = 0;
};

struct __declspec(uuid("8b217763-717d-11ce-ab5b-d41203c10000"))
CustomData : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct CustomData * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall get_GUID (
        BSTR * retVal ) = 0;
    virtual HRESULT __stdcall get_Value (
        VARIANT * retVal ) = 0;
};

struct __declspec(uuid("8b217764-717d-11ce-ab5b-d41203c10000"))
CustomDataCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall Me (
        struct CustomDataCollection * * retVal ) = 0;
    virtual HRESULT __stdcall _placeholder_destructor ( ) = 0;
    virtual HRESULT __stdcall _NewEnum (
        IUnknown * * retVal ) = 0;
    virtual HRESULT __stdcall get_Item (
        short Index,
        struct CustomData * * retVal ) = 0;
    virtual HRESULT __stdcall get_Count (
        short * retVal ) = 0;
};

struct __declspec(uuid("8b217765-717d-11ce-ab5b-d41203c10000"))
UnionInfo : ConstantInfo
{};

struct __declspec(uuid("8b217766-717d-11ce-ab5b-d41203c10000"))
Unions : _BaseTypeInfos
{
    //
    // Raw methods provided by interface
    //

    virtual HRESULT __stdcall get_Item (
        short Index,
        struct UnionInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_IndexedItem (
        short TypeInfoNumber,
        struct UnionInfo * * retVal ) = 0;
    virtual HRESULT __stdcall get_NamedItem (
        BSTR * TypeInfoName,
        struct UnionInfo * * retVal ) = 0;
};

/*
enum CallConvs
{
    CC_FASTCALL = 0,
    CC_CDECL = 1,
    CC_MSCPASCAL = 2,
    CC_PASCAL = 2,
    CC_MACPASCAL = 3,
    CC_STDCALL = 4,
    CC_FPFASTCALL = 5,
    CC_SYSCALL = 6,
    CC_MPWCDECL = 7,
    CC_MPWPASCAL = 8,
    CC_MAX = 9
};
*/

} // namespace TLI

#pragma pack(pop)
