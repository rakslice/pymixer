import argparse
import sys
import session_channel_api

from pycaw.pycaw import AudioUtilities

"""A Windows command line utility to set the per-application mixer volume and pan"""


def parse_args():
    parser = argparse.ArgumentParser(description="A Windows command line utility to set the per-application audio mixer settings")

    # searches

    searches_argument_group = parser.add_argument_group("Application Searches",
                                                        description="These select the applications to show or make volume changes on")

    searches_argument_group.add_argument("--display-name-contains", "-d",
                                         help="Search for PART in the mixer column's display name", metavar="PART")
    searches_argument_group.add_argument("--process", "-P", help="Find the process WHATEVER_EXE",
                                         metavar="WHATEVER_EXE")
    searches_argument_group.add_argument("--pid",
                                         help="Find the process with id PID",
                                         metavar="PID",
                                         type=int)
    searches_argument_group.add_argument("--index", "-i",
                                         help="Use the Ith matching process, numbered starting at 0",
                                         metavar="I",
                                         type=int)

    # actions
    actions_argument_group = parser.add_argument_group("Actions",
                                                       description="These are actions to take on the matching applications")

    actions_argument_group.add_argument("--list", "-l",
                                        help="Show process info and current audio settings",
                                        default=False,
                                        action="store_true")
    actions_argument_group.add_argument("--pan", "-p",
                                        help="Pan the audio to POS (-1.0 [left] to 1.0 [right])",
                                        type=float,
                                        metavar="POS")
    actions_argument_group.add_argument("--left",
                                        help="Pan the audio left",
                                        default=False,
                                        action="store_true")
    actions_argument_group.add_argument("--right",
                                        help="Pan the audio left",
                                        default=False,
                                        action="store_true")
    actions_argument_group.add_argument("--center",
                                        help="Pan the audio to the center",
                                        default=False,
                                        action="store_true")
    actions_argument_group.add_argument("--vol", "-v",
                                        help="Set the audio level to VOL (0.0 to 1.0)",
                                        type=float,
                                        metavar="VOL")
    actions_argument_group.add_argument("--mute", "-m",
                                        help="Mute the app",
                                        default=False,
                                        action="store_true")
    actions_argument_group.add_argument("--unmute", "-u",
                                        help="Unmute the app",
                                        default=False,
                                        action="store_true")

    # toggles

    toggles_argument_group = parser.add_argument_group("Toggles")
    toggles_argument_group.add_argument("--single", "-s",
                                        help="If no apps or multiple apps   match, treat it as an error (by default adjust whatever matches)",
                                        default=False,
                                        action="store_true",
                                        )
    toggles_argument_group.add_argument("--quiet", "-q",
                                        help="Don't output any messages about adjustment actions",
                                        default=False,
                                        action="store_true",
                                        )
    return parser.parse_args()


def log_stderr(msg):
    print >> sys.stderr, msg
    sys.stderr.flush()


def die(msg):
    log_stderr(msg)
    sys.exit(1)


def convert_pan_to_channel_levels(pan):
    """
    Convert the pan to channel level values
    :type pan: float
    :param pan: [-1.0 .. 1.0]
    """
    # 1.0 =>  (0, 1)
    # 0.5 =>  (0.5, 1)
    # 0.20 => (0.8, 1)
    # 0.0 =>  (1, 1)
    # -0.5 => (1, 0.5)
    # -1.0 => (1, 0)

    other_channel_level = 1.0 - abs(pan)
    if pan < 0:
        return 1.0, other_channel_level
    else:
        return other_channel_level, 1.0


def main():
    options = parse_args()

    # validate the options

    if [options.pan is not None, options.left, options.right, options.center].count(True) > 1:
        die("Choose at most one of pan, left, right, and center actions")

    if options.mute and options.unmute:
        die("Choose at most one of mute and unmute actions")

    pan_value = None
    if options.pan is not None:
        if options.pan < -1.0 or options.pan > 1.0:
            die("Pan value not in the range -1.0 to 1.0")
        pan_value = options.pan
    if options.left:
        pan_value = -1.0
    if options.right:
        pan_value = 1.0
    if options.center:
        pan_value = 0.0

    if options.vol is not None:
        if options.vol < 0.0 or options.vol > 1.0:
            die("Volume value not in the range 0.0 to 1.0")

    do_list = options.list

    if not any([options.list, options.pan is not None, options.vol is not None, options.mute, options.unmute,
                options.left, options.right, options.center]):
        log_stderr("No action selected; listing")
        do_list = True

    # find sessions to adjust the volume on

    matching_sessions = []
    # devices = AudioUtilities.GetSpeakers()
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        cur_session_matches = True

        if options.display_name_contains is not None:
            cur_display_name = session.DisplayName
            if options.display_name_contains not in cur_display_name:
                cur_session_matches = False

        if options.process is not None:
            if session.Process is None:
                cur_process_name = ""
            else:
                cur_process_name = session.Process.name()
            if cur_process_name.lower() != options.process.lower():
                cur_session_matches = False

        if options.pid is not None:
            if session.Process is None:
                cur_pid = None
            else:
                cur_pid = session.Process.pid
            if cur_pid != options.pid:
                cur_session_matches = False

        if cur_session_matches:
            matching_sessions.append(session)

    # do any needed validation of the matched sessions list

    if options.single:
        if len(matching_sessions) == 0:
            die("No matching sessions found")
        if len(matching_sessions) > 1:
            die("Multiple matching sessions found")

    if options.index is not None:
        if 0 <= options.index < len(matching_sessions):
            indexed_session = matching_sessions[options.index]
            matching_sessions = [indexed_session]
        else:
            die("No matching session #%d" % options.index)

    # take any specified actions on the matched sessions

    for session_index, session in enumerate(matching_sessions):
        simple_audio_volume = session.SimpleAudioVolume

        channels_api = session_channel_api.get_for_session(session)

        num_channels = channels_api.GetChannelCount()

        if options.index is not None:
            session_index = options.index
        if not options.quiet:
            print (u'#%d "%s"' % (session_index, session.DisplayName)).encode("utf-8")

        if do_list:
            cur_process_name = None
            if session.Process is not None:
                cur_process_name = session.Process.name()
            print "Process: %s" % cur_process_name
            print "Process Id: %d" % session.ProcessId

            print "Current audio settings:"
            cur_muted = simple_audio_volume.GetMute()
            print "  Muted: %s" % ("Yes" if cur_muted else "No")
            print "  Volume: %0.2f" % simple_audio_volume.GetMasterVolume()
            print "  Channels: %d" % num_channels

            for i in range(num_channels):
                channel_volume = channels_api.GetChannelVolume(i)
                print "    Channel %d Level: %0.2f" % (i, channel_volume)

            print ""

        if pan_value is not None:
            if num_channels != 2:
                log_stderr("Session %s doesn't have 2 channels, so can't pan. Skipping it.")
            else:
                left_channel_level, right_channel_level = convert_pan_to_channel_levels(pan_value)
                if not options.quiet:
                    pos_names = {0.0: "center", -1.0: "left", 1.0: "right"}
                    if pan_value in pos_names:
                        pos_name = "%s (%s)" % (pos_names[pan_value], pan_value)
                    else:
                        pos_name = pan_value
                    log_stderr("Panning to %s [%s, %s]" % (pos_name, left_channel_level, right_channel_level))
                channels_api.SetChannelVolume(0, left_channel_level, None)
                channels_api.SetChannelVolume(1, right_channel_level, None)

        if options.vol is not None:
            if not options.quiet:
                log_stderr("Setting volume to %s" % options.vol)
            simple_audio_volume.SetMasterVolume(options.vol, None)

        if options.mute:
            if not options.quiet:
                log_stderr("Muting")
            simple_audio_volume.SetMute(True, None)

        if options.unmute:
            if not options.quiet:
                log_stderr("Unmuting")
            simple_audio_volume.SetMute(False, None)


if __name__ == "__main__":
    main()
