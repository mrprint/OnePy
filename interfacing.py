# -*- coding: utf-8 -*-

__all__ = [
    'VERSION',
    'TARGETPLATFORM',
    'EPLATFORMCLASS',  # 'DEBUG',
    'IB5826CA69712448A99067B6F1938F6B2',
    'D0C4EF5E1030946E0B35458F15D7B79D3',
#$if $$notdemo
    'P5793C5CF34DB467985C5D4952D538DBE',
#$end if
    'ContainerClasses',
    'CreateNewObject',
    'ToValueTable',
    'EntObjectProxy',
    'AppDispatchProxy',
    'UseAppDispatch',
    'UseNewObjects',
    'UseAppDispatchProxy',
    'UseNewObjectsProxy',  # 'UsesEnterprise',
    'Component',
    'Readable',
    'Writable',
    'Procedure',
    'Function'
]

import clr
import inspect, new, threading
from functools import wraps
from itertools import izip
from System import Array, Object
from System.Reflection import BindingFlags, MemberTypes
from System.Runtime.InteropServices import Marshal

import logging

#$if $$notdemo
VERSION = '1.0'
#$else
VERSION = 'demo'
#$end if
TARGETPLATFORM = 'DOTNET$$TargetId:$$PlatformName'
#$if $$notdemo
EPLATFORMCLASS = None
#$else
EPLATFORMCLASS = 7
#$end if

#$if $$debug
# DEBUG = True
#$else
# DEBUG = False
#$end if
IB5826CA69712448A99067B6F1938F6B2 = None
D0C4EF5E1030946E0B35458F15D7B79D3 = {'P': {}, 'M': []}

TypeAliases = {
    u'КОНСТАНТА': 'CONST',
    u'СПРАВОЧНИК': 'REFERENCE',
    u'ПЕРЕЧИСЛЕНИЕ': 'ENUM',
    u'ДОКУМЕНТ': 'DOCUMENT',
    u'РЕГИСТР': 'REGISTER',
    u'ПЛАНСЧЕТОВ': 'CHARTOFACCOUNTS',
    u'СЧЕТ': 'ACCOUNT',
    u'ВИДСУБКОНТО': 'SUBCONTOKIND',
    u'ОПЕРАЦИЯ': 'OPERATION',
    u'БУХГАЛТЕРСКИЕИТОГИ': 'BOOKKEEPINGTOTALS',
    u'ЖУРНАЛРАСЧЕТОВ': 'CALCJOURNAL',
    u'ВИДРАСЧЕТА': 'CALCULATIONKIND',
    u'ГРУППАРАСЧЕТОВ': 'CALCULATIONGROUP',
    u'КАЛЕНДАРЬ': 'CALENDAR',
    u'ЗАПРОС': 'QUERY',
    u'ТЕКСТ': 'TEXT',
    u'ТАБЛИЦА': 'TABLE',
    u'СПИСОКЗНАЧЕНИЙ': 'VALUELIST',
    u'ТАБЛИЦАЗНАЧЕНИЙ': 'VALUETABLE',
    u'КАРТИНКА': 'PICTURE',
    u'ПЕРИОДИЧЕСКИЙ': 'PЕRIODIC',
    u'ФС': 'FS'
}


def BaseType(name):
    # try:
    id_ = name.split('.')[0].upper()
    return TypeAliases.get(id_, id_)
    # except:
    #    pass
    # return None


def GetClass(name):
    bt = BaseType(name)
    if bt in ContainerClasses:
        return ContainerClasses[bt]
    return EntObjectProxy


# эвристическое создание новых объектов

#$if $$notdemo
def __CreateDefault(disp, name):
    global __Creator
    result = None
    __Creator = None  # если что-то пойдёт не так, впоследствии возникнут характерные исключения
    if EPLATFORMCLASS == 7:
        result = __CreateOldStyle(disp, name)
        __Creator = __CreateOldStyle
        return result
    elif EPLATFORMCLASS == 8:
        result = __CreateNewStyle(disp, name)
        __Creator = __CreateNewStyle
        return result


def __CreateNewStyle(disp, name):
    return clr.GetClrType(type(disp)).InvokeMember('NewObject',
                                                   BindingFlags.Public | BindingFlags.InvokeMethod | BindingFlags.Static | BindingFlags.CreateInstance,
                                                   None, disp, Array[object]([name]))


#$end if

def __CreateOldStyle(disp, name):
    return clr.GetClrType(type(disp)).InvokeMember('CreateObject', BindingFlags.Public | BindingFlags.GetProperty, None,
                                                   disp, Array[object]([name]))


def CreateNewObject(disp, name):
    return __Creator(disp, name)


#$if $$notdemo
__Creator = __CreateDefault
#$else
__Creator = __CreateOldStyle


#$end if

# трансформирование коллекции в таблицу значений

def ToValueTable(ent, vt, it):
    first = True
    for dr in it:
        if first:
            wd = 1
            rit = True
            try:
                wd = len(it[0])
            except TypeError, e:
                rit = False
            if isinstance(vt, EntValueTableProxy):
                vtp = vt
            else:
                vtp = EntValueTableProxy(ent, vt)
            vtp.Clear()
            for i in xrange(1, wd + 1):
                vtp.NewColumn(u'Н%s' % i)
            first = False
        vtp.append(dr if rit else [dr])


# протоколы контейнеров и итерации для типов 1С

class ContainerReference(object):
    def Init(self):
        self.SelectItems()

    @property
    def Value(self):
        return self.CurrentItem()

    def __iter__(self):
        self.Init()
        return self

    def next(self):
        if self.GetItem() != 1:
            raise StopIteration
        return self

    def __reversed__(self):
        self.BackwardOrder()
        return self

    def append(self, val):
        pass


class ContainerValueTable(object):
    class ContainersRow(object):

        def __init__(self, base, rowi=-1):
            self.base = base
            self.rowi = rowi

        def __len__(self):
            return self.self.base.__getattr__(u'КоличествоКолонок')()

        def __iter__(self):
            self.ccol = 0
            return self

        def next(self):
            if self.ccol >= len(self):
                raise StopIteration
            v = self.__getitem__(self.ccol)
            self.ccol += 1
            return v

        def __getattr__(self, name):
            if self.rowi > -1:
                return self.base.__getattr__(u'ПолучитьСтрокуПоНомеру')(self.rowi + 1)
            return self.base.__getattr__(name)

        def __getitem__(self, key):
            if isinstance(key, int):
                key += 1
            if self.rowi > -1:
                ri = self.rowi + 1
            else:
                ri = self.base.property.__getattr__(u'НомерСтроки')
            return self.base.__getattr__(u'ПолучитьЗначение')(ri, key)

        def __setitem__(self, key, value):
            if isinstance(key, int):
                key += 1
            if self.rowi > -1:
                ri = self.rowi + 1
            else:
                ri = self.base.property.__getattr__(u'НомерСтроки')
            self.base.__getattr__(u'УстановитьЗначение')(ri, key, value)

    def Init(self):
        self.SelectLines()

    @property
    def Value(self):
        li = self.CurrentLine()
        return tuple([self.GetValue(li, ci) for ci in xrange(1, self.ColumnCount() + 1)])

    def __len__(self):
        return self.LinesCnt()

    def __iter__(self):
        self.Init()
        return self

    def next(self):
        if self.GetLine() != 1:
            raise StopIteration
        return self.ContainersRow(self)

    def __getitem__(self, key):
        return self.ContainersRow(self, key)

    def append(self, val):
        li = self.LinesCnt() + 1
        self.NewLine(li)
        for i, v in enumerate(val, start=1):
            self.SetValue(li, i, v)


ContainerClasses = {}


# классы поддержки взаимодействия COM объектов

class EntObjectProxy(object):
    """Обёртка объекта 1С."""

    class Field(object):

        def __init__(self, base):
            self.__dict__['base'] = base

        def __getattr__(self, name):
            return self.base.adt.InvokeMember(name, BindingFlags.Public | BindingFlags.GetField, None, self.base.obj,
                                              None)

        def __setattr__(self, name, val):
            try:
                self.base.adt.InvokeMember(name, BindingFlags.Public | BindingFlags.SetField, None, self.base.obj,
                                           Array[object]([val]))
            except Exception, e:
                try:
                    super(Field, self).__setattr__(name, value)
                except:
                    raise e

    class Property(object):

        def __init__(self, base):
            self.__dict__['base'] = base

        def __getattr__(self, name):
            return self.base.adt.InvokeMember(name, BindingFlags.Public | BindingFlags.GetProperty, None, self.base.obj,
                                              None)

        def __setattr__(self, name, val):
            try:
                self.base.adt.InvokeMember(name, BindingFlags.Public | BindingFlags.PutDispProperty, None,
                                           self.base.obj, Array[object]([val]))
            except Exception, e:
                try:
                    super(Property, self).__setattr__(name, value)
                except:
                    raise e

    def __init__(self, ent, obj, name=None):
        self.obj = obj
        self.name = name
        with UseAppDispatch(ent) as ad:
            self.adt = clr.GetClrType(type(ad))

    def __getattr__(self, name):
        uname = name.upper()
        if uname in ['PROPERTY', u'СВОЙСТВО']:
            return self.Property(self)
        elif uname in ['FIELD', u'ПОЛЕ']:
            return self.Field(self)

        def mbody(*args, **kwargs):
            iattr = BindingFlags.Public | BindingFlags.InvokeMethod
            if 'invokeAttr' in kwargs:
                iattr = kwargs['invokeAttr']
            return self.adt.InvokeMember(name, iattr, None, self.obj, Array[object](args))

        return mbody

    @property
    def val(self):
        return self.obj


class AppDispatchProxy(EntObjectProxy):
    def __init__(self, obj):
        self.obj = obj
        # if EPLATFORMCLASS == 7:
        # self.adt = self.obj.GetType()
        self.adt = clr.GetClrType(type(self.obj))

    def __getattr__(self, name):
        uname = name.upper()

        def nbody(*args, **kwargs):
            return CreateNewObject(self.obj, args[0])

        if uname in ['CREATEOBJECT', 'NEWOBJECT', u'СОЗДАТЬОБЪЕКТ', u'НОВЫЙ']:
            return nbody
        return super(AppDispatchProxy, self).__getattr__(name)


# менеджеры контекста

class UseAppDispatch(object):
    """Менеджер контекста AppDispatch."""

    counter = 0

    def __init__(self, ent):
        self.ent = ent
        self.ad = None

    def __enter__(self):
#$if $$notdemo
        global EPLATFORMCLASS
        with threading.RLock():  # Без этого не работает. Нужно разбираться.
            self.ad = self.ent.like_dispatch.GetType().InvokeMember('AppDispatch',
                                                                    BindingFlags.Public | BindingFlags.GetProperty,
                                                                    None, self.ent.like_dispatch, None)
        try:
            self.ad.GetType()
        except:
            EPLATFORMCLASS = 8
        else:
            EPLATFORMCLASS = 7
            Marshal.GetIUnknownForObject(self.ad)
#$else
        with threading.RLock():  # Без этого не работает. Нужно разбираться.
            # self.ad = self.ent.like_dispatch.GetType().InvokeMember('AppDispatch', BindingFlags.Public | BindingFlags.InvokeMethod, None, self.ent.like_dispatch, None)
            self.ad = self.ent.like_dispatch.GetType().InvokeMember('AppDispatch',
                                                                    BindingFlags.Public | BindingFlags.GetProperty,
                                                                    None, self.ent.like_dispatch, None)
        Marshal.GetIUnknownForObject(self.ad)
#$end if
        UseAppDispatch.counter += 1
        return self.ad

    def __exit__(self, exc_type, exc_val, traceback):
        Marshal.Release(Marshal.GetIDispatchForObject(self.ad))
        UseAppDispatch.counter -= 1
        if UseAppDispatch.counter == 0:
            Marshal.FinalReleaseComObject(self.ad)


class UseNewObjects(object):
    """Менеджер контекста COM объектов 1С."""

    def __init__(self, ent, *names):
        self.ent = ent
        self.names = names
        self.objs = []

    def __enter__(self):
        """Возвращает tuple для нескольких объектов, или единственный."""

        with UseAppDispatch(self.ent) as ad:
            # adt = ad.GetType()
            for n in self.names:
                self.objs.append(CreateNewObject(ad, n))
                # self.objs.append(adt.InvokeMember('CreateObject', BindingFlags.GetProperty, None, ad, Array[object]([n])))
        if len(self.objs) == 1:
            return self.objs[0]
        return tuple(self.objs)

    def __exit__(self, exc_type, exc_val, traceback):
        for obj in self.objs:
            Marshal.FinalReleaseComObject(obj)


class UseAppDispatchProxy(UseAppDispatch):
    def __enter__(self):
        ad = super(UseAppDispatchProxy, self).__enter__()
        return AppDispatchProxy(ad)


class UseNewObjectsProxy(UseNewObjects):
    def __enter__(self):
        vals = super(UseNewObjectsProxy, self).__enter__()
        if isinstance(vals, tuple):
            return tuple([GetClass(nm)(self.ent, ob) for ob, nm in izip(vals, self.names)])
        return GetClass(self.names[0])(self.ent, vals)


# узнаём тип платформы

#$if $$notdemo
def P5793C5CF34DB467985C5D4952D538DBE(ent):
    with UseAppDispatch(ent) as ad:
        pass
#$end if

# Классы обёрток агрегатных

for id, bcl in [('Reference', ContainerReference),
                ('ValueTable', ContainerValueTable)]:
    clname = 'Ent%sProxy' % id
    cl = new.classobj(clname, (EntObjectProxy, bcl), {})
    globals()[clname] = cl
    ContainerClasses[id.upper()] = cl
    __all__.append(clname)


# определение интерфейса с модулями        

class UsesEnterprise(object):
    def __init__(self, inter):
        self.Enterprise = inter

    def ExplicitDone(self):
        pass


def Component(cls):
    global IB5826CA69712448A99067B6F1938F6B2

    class MainIntClass(cls, UsesEnterprise):
        def __init__(self, inter):
            UsesEnterprise.__init__(self, inter)
            if hasattr(cls, '__init__'):
                cls.__init__(self)

        def ExplicitDone(self):
            if hasattr(cls, 'ExplicitDone'):
                cls.ExplicitDone(self)
            UsesEnterprise.ExplicitDone(self)

    oldname = cls.__name__
    MainIntClass.__name__ = 'MainIntClass_' + oldname
    IB5826CA69712448A99067B6F1938F6B2 = MainIntClass
    return MainIntClass


# предшествует стандартной @property
def Readable(*names):
    global D0C4EF5E1030946E0B35458F15D7B79D3

    def dcr(prp):
        assert isinstance(prp, property)
        D0C4EF5E1030946E0B35458F15D7B79D3['P'][prp.fget] = [names, None]
        return prp

    return dcr


def Writable(prp):
    global D0C4EF5E1030946E0B35458F15D7B79D3
    assert isinstance(prp, property)
    D0C4EF5E1030946E0B35458F15D7B79D3['P'][prp.fget][1] = prp.fset
    return prp


def Procedure(*names):
    global D0C4EF5E1030946E0B35458F15D7B79D3

    def dcr(func):
        ai = inspect.getargspec(func)
        D0C4EF5E1030946E0B35458F15D7B79D3['M'].append(
            (names, len(ai.args) - 1, ai.defaults if ai.defaults else [], False, func))
        return func

    return dcr


def Function(*names):
    global D0C4EF5E1030946E0B35458F15D7B79D3

    def dcr(func):
        ai = inspect.getargspec(func)
        D0C4EF5E1030946E0B35458F15D7B79D3['M'].append(
            (names, len(ai.args) - 1, ai.defaults if ai.defaults else [], True, func))
        return func

    return dcr
