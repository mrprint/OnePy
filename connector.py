# -*- coding: utf-8 -*-

__all__ = ['OnePyConnector']

# import locale
# locale.setlocale(locale.LC_ALL, '')

import sys
import os
import gc
import weakref
import logging
from collections import namedtuple

logging.basicConfig(format=u'%(levelname)-8s [%(asctime)s] %(message)s', level=logging.INFO,
                    filename=os.path.expandvars(u'${APPDATA}\\onepy\\python.log'))
#$if not $$debug
# logging.disable(logging.ERROR)
#$end if

MethInfoRec = namedtuple('MethInfoRec', 'names argc dfltc isfunc inst method')
PropInfoRec = namedtuple('PropInfoRec', 'names setr getr inst')


# ContextInfoRec = namedtuple('ContextInfoRec', 'globals0 main_module1 main_class2 main_object3 methods4 dmethods5 properties6 dproperties7')

class ScopeInfo(object):
    def __init__(self):
        self.clean()

    def clean(self):
        self.globals = {}
        # self.main_module = None
        self.main_class = None
        # del self.main_object
        self.main_object = None
        self.methods = None
        self.dmethods = None
        self.properties = None
        self.dproperties = None


class ModulesInfo(object):
    def __init__(self, mc, mt, pr):
        self.mclass, self.methodsp, self.propertiesp = mc, mt, pr

    def clean(self):
        self.mclass = None
        self.methodsp = None
        self.propertiesp = None


# class Output(object):
#    
#    def __init__(self, inter, sign):
#        self.interactor = inter
#        self.sign = sign
#        
#    def write(self, s):
#        ei = self.interactor.EXCEPINFOProxy()
#        ei.wCode = self.sign
#        ei.bstrDescription = s
#        ei.bstrSource = ''
#        self.interactor.Message(ei)

class Interactor(object):
    def __init__(self, inter):
        self.data = inter

    @property
    def like_dispatch(self):
        return self.data.like_dispatch

    @property
    def like_asyncevent(self):
        return self.data.like_asyncevent

    @property
    def like_errorlog(self):
        return self.data.like_errorlog

    @property
    def like_statusline(self):
        return self.data.like_statusline

    @property
    def like_extwndssupport(self):
        return self.data.like_extwndssupport

    @property
    def like_propertyprofile(self):
        return self.data.like_propertyprofile

    @property
    def Message(self):
        return self.data.Message

    @property
    def EXCEPINFOProxy(self):
        return self.data.EXCEPINFOCover


class OnePyConnector(object):
    def __init__(self, inter):
        self.__interactor = Interactor(inter)
        self.interactor = weakref.proxy(self.__interactor)
        self.modules_info = {}
        self.scopes = {}
        self.__SwitchScope()

    #        self.__Configure()

    def GetNProps(self, plProps):
        return len(self.current_scope.properties)

    def FindProp(self, bstrPropName, plPropNum):
        return self.current_scope.dproperties[bstrPropName][0]

    def GetPropName(self, lPropNum, lPropAlias, pbstrPropName):
        pass

    def GetPropVal(self, lPropNum, pvarPropVal):
        d = self.current_scope.properties[lPropNum]
        return d.getr(d.inst)

    def SetPropVal(self, lPropNum, varPropVal):
        d = self.current_scope.properties[lPropNum]
        return d.setr(d.inst, varPropVal)

    def IsPropReadable(self, lPropNum, pboolPropRead):
        return True

    def IsPropWritable(self, lPropNum, pboolPropWrite):
        return self.current_scope.properties[lPropNum].setr is not None

    def GetNMethods(self, plMethods):
        return len(self.current_scope.methods)  # plMethods

    def FindMethod(self, bstrMethodName, plMethodNum):
        return self.current_scope.dmethods[bstrMethodName][0]

    def GetMethodName(self, lMethodNum, lMethodAlias, pbstrMethodName):
        pass

    def GetNParams(self, lMethodNum, plParams):
        return self.current_scope.methods[lMethodNum].argc

    def GetParamDefValue(self, lMethodNum, lParamNum, pvarParamDefValue):
        md = self.current_scope.methods[lMethodNum]
        dl = md.dfltc
        i = lParamNum - (md.argc - len(dl))
        if i >= 0:
            return dl[i]
        return pvarParamDefValue

    def HasRetVal(self, lMethodNum, pboolRetValue):
        return self.current_scope.methods[lMethodNum].isfunc

    def CallAsProc(self, lMethodNum, pParams):
        args = [self.current_scope.methods[lMethodNum].inst]
        if pParams:
            args += list(pParams)
        self.current_scope.methods[lMethodNum].method(*args)
        return 1

    def CallAsFunc(self, lMethodNum, pvarRetValue, pParams):
        args = [self.current_scope.methods[lMethodNum].inst]
        if pParams:
            args += list(pParams)
        return self.current_scope.methods[lMethodNum].method(*args)

    # !оптимизировать    
    def CleanAll(self):
        for id_ in self.scopes:
            self.current_scope = self.scopes[id_]
            if self.current_scope.main_class:
                self.current_scope.main_class.ExplicitDone(self.current_scope.main_object)
            self.current_scope.clean()
        sys.exc_clear()
        gc.collect()
        del gc.garbage[:]
        self.__SwitchScope()
        self.__FlushMethods()
        self.__CollectModuleInfo()

    def AddReferences(self, ls):
        import clr
        for i in ls:
            if i:
                clr.AddReference(i)

    def ClearExcInfo(self):
        sys.exc_clear()

    #    def __Configure(self):
    #        from interfacing import UseAppDispatch

    #    def AfterInit(self):
    #        sys.stdout = Output(1001)
    #        sys.stderr = Output(1006)

    def __FlushMethods(self):
        # имена, количество параметров, признак функции, экземпляр, обработчик
        self.current_scope.methods = [
            MethInfoRec((u'ConnectMainModule', u'ПодключитьОсновнойМодуль'), 1, (), True, self,
                        OnePyConnector.__ConnectMainModule),
#$if $$notdemo
            MethInfoRec((u'SwitchScope', u'СменитьКонтекст'), 1, (u'Default',), False, self,
                        OnePyConnector.__SwitchScope),
            MethInfoRec((u'CleanScope', u'ОчиститьКонтекст'), 0, (), False, self, OnePyConnector.__CleanScope),
#$end if
            MethInfoRec((u'CleanAll', u'ОчиститьВсё', u'ОчиститьВсе'), 0, (), False, self, OnePyConnector.CleanAll),
            MethInfoRec((u'PythonSetValue', u'ПитонЗадатьЗначение', u'ПитонПрисвоить'), 2, (), False, self,
                        OnePyConnector.__SetVal),
            MethInfoRec((u'PythonEvaluate', u'ПитонВычислить'), 2, (0,), True, self, OnePyConnector.__Eval),
            MethInfoRec((u'PythonExecute', u'ПитонВыполнить'), 2, (1,), False, self, OnePyConnector.__Exec)]
        self.current_scope.properties = []

    def __CollectModuleInfo(self):
        dmethods = {}
        dproperties = {}
        for c, r in enumerate(self.current_scope.methods):
            for i in r.names:
                if i in dmethods:
                    raise RuntimeError('Duplicate method name "%s"' % i)
                dmethods[i] = (c, r)
        self.current_scope.dmethods = dmethods
        for c, r in enumerate(self.current_scope.properties):
            for i in r.names:
                if i in dproperties:
                    raise RuntimeError('Duplicate property name "%s"' % i)
                dproperties[i] = (c, r)
        self.current_scope.dproperties = dproperties

    def __ConnectMainModule(self, mdl):
        try:
            __import__(mdl, self.current_scope.globals)
            im = sys.modules[mdl]
            # self.current_scope.main_module = im
            self.__SetVal(mdl, im)
            # self.current_scope.globals[mdl] = im
            if mdl in self.modules_info:
                mi = self.modules_info[mdl]
                self.current_scope.main_class = mi.mclass
                self.current_scope.main_object = mi.mclass(self.interactor)
                self.__FlushMethods()
                self.current_scope.methods += mi.methodsp
                self.current_scope.properties = mi.propertiesp
                self.__CollectModuleInfo()
            else:
                from interfacing import IB5826CA69712448A99067B6F1938F6B2 as MClass, \
                    D0C4EF5E1030946E0B35458F15D7B79D3 as collected
                self.current_scope.main_class = MClass
                self.current_scope.main_object = MClass(self.interactor)
                self.__FlushMethods()
                methodsp = [MethInfoRec(n, c, d, f, self.current_scope.main_object, m) for n, c, d, f, m in
                            collected['M']]
                self.current_scope.methods += methodsp
                prl = collected['P']
                propertiesp = [PropInfoRec(*v + [k, self.current_scope.main_object]) for k, v in prl.iteritems()]
                self.current_scope.properties = propertiesp
                self.__CollectModuleInfo()
                self.modules_info[mdl] = ModulesInfo(MClass, methodsp, propertiesp)
            return 0
        except Exception, e:
            logging.error(e)
        return 1

    def __SwitchScope(self, id_='Default'):
        if id_ == u'ПоУмолчанию':
            id_ = 'Default'
        if id_ in self.scopes:
            self.current_scope = self.scopes[id_]
        else:
            newc = ScopeInfo()
            self.scopes[id_] = newc
            self.current_scope = newc
            self.__FlushMethods()
            self.__CollectModuleInfo()
            self.__SetVal('Enterprise', self.interactor)
#$if $$notdemo
            self.current_scope.globals.update(
                __import__('interfacing', self.current_scope.globals, fromlist=['P5793C5CF34DB467985C5D4952D538DBE']).__dict__)
            self.current_scope.globals['P5793C5CF34DB467985C5D4952D538DBE'](self.interactor)
#$end if
            self.current_scope.globals.update(
                __import__('interfacing', self.current_scope.globals, fromlist=['*']).__dict__)

            #exec co_interf in self.current_scope.globals

    def __SetVal(self, var, vl):
        self.current_scope.globals[var.strip()] = vl

    def __Eval(self, expr, comp=0):
        if bool(comp):
            return eval(compile(expr.translate(None, '\r'), '<string>', 'eval'), self.current_scope.globals)
        else:
            return eval(expr.translate(None, '\r'), self.current_scope.globals)

    def __Exec(self, expr, comp=1):
        if bool(comp):
            exec compile(expr.translate(None, '\r'), '<string>', 'exec') in self.current_scope.globals
        else:
            exec expr.translate(None, '\r') in self.current_scope.globals

    def __CleanScope(self):
        if self.current_scope.main_class:
            self.current_scope.main_class.ExplicitDone(self.current_scope.main_object)
        self.current_scope.clean()
        sys.exc_clear()
        gc.collect()
        del gc.garbage[:]
        self.__FlushMethods()
        self.__CollectModuleInfo()
