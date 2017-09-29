
""" Wrapper for IChannelAudioVolume """

# TODO push all this up to pycaw

import pycaw.pycaw

from ctypes import HRESULT, POINTER, c_uint32, c_float
from comtypes import IUnknown, GUID, COMMETHOD


def get_for_session(session):
    assert isinstance(session, pycaw.pycaw.AudioSession)
    # noinspection PyProtectedMember
    return session._ctl.QueryInterface(IChannelAudioVolume)


class IChannelAudioVolume(IUnknown):
    _iid_ = GUID('{1c158861-b533-4b30-b1cf-e853e51c59b8}')
    # GUID courtesy of https://github.com/wine-mirror/wine/blob/master/include/audioclient.idl
    _methods_ = (
        # HRESULT GetChannelCount(
        # [out] UINT32 *pdwCount
        # );
        COMMETHOD([], HRESULT, 'GetChannelCount',
                  (['out'], POINTER(c_uint32), 'pdwCount')),

        # HRESULT SetChannelVolume(
        # [in] UINT32 dwIndex,
        # [in] const float fLevel,
        # [unique,in] LPCGUID EventContext
        # );
        COMMETHOD([], HRESULT, 'SetChannelVolume',
                  (['in'], c_uint32, 'dwIndex'),
                  (['in'], c_float, 'fLevel'),
                  (['in'], POINTER(GUID), 'pguidEventContext')),

        # HRESULT GetChannelVolume(
        #     [in] UINT32 dwIndex,
        #                 [out] float *pfLevel
        # );
        COMMETHOD([], HRESULT, 'GetChannelVolume',
                  (['in'], c_uint32, 'dwIndex'),
                  (['out'], POINTER(c_float))),

        # HRESULT SetAllVolumes(
        #     [in] UINT32 dwCount,
        #                 [size_is(dwCount),in] const float *pfVolumes,
        #                                             [unique,in] LPCGUID EventContext
        # );
        # Not implementing yet

        # HRESULT GetAllVolumes(
        #     [in] UINT32 dwCount,
        #                 [size_is(dwCount),out] float *pfVolumes
        # );
        # Not implementing yet
    )
