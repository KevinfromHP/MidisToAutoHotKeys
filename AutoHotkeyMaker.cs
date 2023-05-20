using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using NAudio.Midi;
using System.Linq;

namespace AutoHotkeyMaker
{
    class Program
    {
        public static List<MidiFile> midiFiles = new List<MidiFile>();
        public static List<String> midiNames = new List<string>();
        public static List<String> hotKeys = new List<string>();
        public static int transpose;
        public static int songDeltaTicks;
        public static float currentTempo;
        public static int songCount = 0;
        public static List<Char> thingsAdded = new List<Char>();
        public static string keys = "1234567890qwertyuiopasdfghjklzxcvbnm";
        public static string[] keyValues = { "vk31sc002", "vk32sc003", "vk33sc004", "vk34sc005", "vk35sc006", "vk36sc007", "vk37sc008", "vk38sc009", "vk39sc00A", "vk30sc00B", "vk51sc018", "vk57sc011", "vk45sc02E", "vk52sc013", "vk54sc014", "vk59sc015", "vk55sc016", "vk49sc022", "vk4Fsc032", "vk50sc031", "vk41sc01E", "vk53sc01F", "vk44sc01D", "vk46sc020", "vk47sc012", "vk48sc021", "vk4Asc023", "vk4Bsc017", "vk4Csc024", "vk5Asc02c", "vk58sc02D", "vk43sc02F", "vk56sc02F", "vk42sc030", "vk4Esc026", "vk4Dsc025" };

        public static string directory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);


        public static void Main(string[] args)
        {
            Console.WriteLine("Thank you for using my script!");
            Console.WriteLine("To use this generator, please open it with one or more .mid files. This generator only excepts .mid files!");
            bool hasArgs = true;
            foreach (String thing in args)
                if (!thing.EndsWith(".mid"))
                    hasArgs = false;
            if (hasArgs && args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    string midiDirectory = args[i];
                    Console.WriteLine(midiDirectory);
                    midiNames.Add(midiDirectory.Substring(midiDirectory.LastIndexOf('\\') + 1));
                    Console.WriteLine("Please assign a hotkey to " + midiNames[i] + ". If you are not sure what to type in for your start key, refer to www.autohotkey.com/docs/KeyList.htm");
                    hotKeys.Add(Console.ReadLine());
                    Console.WriteLine("Type in the number of octaves you would like to transpose by. If none, type 0.");
                    transpose = Int32.Parse(Console.ReadLine());
                    midiFiles.Add(new MidiFile(File.OpenRead(midiDirectory), false));
                    ParseNotes(midiFiles[i]);
                    Console.WriteLine("Autohotkey written.\n");
                    songCount++;
                }
                Environment.Exit(0);
            }
        }

        static void ParseNotes(MidiFile midi)
        {
            songDeltaTicks = midi.DeltaTicksPerQuarterNote;
            var eventTracks = midi.Events;
            List<MidiEvent> sequenceOfEvents = new List<MidiEvent>();
            int eventTotal = 0;
            for (int i = 0; i < eventTracks.Tracks; i++)
                eventTotal += eventTracks[i].Count - 1;
            int[] eventTrackEventIndex = new int[eventTracks.Tracks]; // These mark the event that each track has progressed through in the for-loop below
            for (int i = 0; i < eventTotal; i++)
            {
                int nextTrack = 0;
                long soonestTime = int.MaxValue;
                for (int k = 0; k < eventTracks.Tracks; k++)
                {
                    //"if this event track's next event plays sooner than the one currently set, set it to be the one that plays next
                    if (eventTrackEventIndex[k] < eventTracks[k].Count && eventTracks[k][eventTrackEventIndex[k]].AbsoluteTime < soonestTime)
                    {
                        nextTrack = k;
                        soonestTime = eventTracks[k][eventTrackEventIndex[k]].AbsoluteTime;
                    }
                }
                sequenceOfEvents.Add(eventTracks[nextTrack][eventTrackEventIndex[nextTrack]]);
                eventTrackEventIndex[nextTrack] += 1;
            }

            //Now we have a list of every event in sequential order in a single array. We need the ones that only affect notes and Tempo changes.

            List<MidiEvent> noteEvents = new List<MidiEvent>();
            foreach (MidiEvent midiEvent in sequenceOfEvents)
            {
                if (midiEvent is NoteEvent || midiEvent is TempoEvent)
                    noteEvents.Add(midiEvent);
            }

            //Now we can start writing the hotkey
            List<String> toBeWritten = new List<String>();
            toBeWritten.Add("#SingleInstance force");
            toBeWritten.Add(hotKeys[songCount] + "::PlaySong()");
            toBeWritten.Add("^RCtrl::ExitApp");
            toBeWritten.Add("RCtrl::Pause");
            toBeWritten.Add("");
            toBeWritten.Add("PlaySong()\n{");
            WriteKeys(noteEvents, toBeWritten);
        }

        static void WriteKeys(List<MidiEvent> events, List<String> lines)
        {
            //Finds start tempo. This is only here just in case it doesn't catch it on the first event.
            currentTempo = (float)events.OfType<TempoEvent>().First().Tempo;

            var orderedEvents = new List<List<MidiEvent>>();
            var sortedEvents = events.GroupBy(mEvent => (float)mEvent.AbsoluteTime).OrderBy(group => group.Key).Select(group => group.ToList()).ToList();

            orderedEvents.Add(sortedEvents[0]);
            for (int i = 1; i < sortedEvents.Count(); i++)
            {
                if (!(sortedEvents[i].OfType<TempoEvent>().Any() || sortedEvents[i].OfType<NoteEvent>().Any()) || (sortedEvents[i].OfType<TempoEvent>().Any() && !sortedEvents[i].OfType<NoteEvent>().Any()))
                    continue;
                float current = sortedEvents[i].First().AbsoluteTime;
                float previous = sortedEvents[i - 1].Last().AbsoluteTime;

                if (current - previous <= 1)
                    orderedEvents[orderedEvents.Count - 1].AddRange(sortedEvents[i]);
                else
                    orderedEvents.Add(sortedEvents[i]);
            }

            float lastTime = 0;



            foreach (var eventSet in orderedEvents)
            {
                float time = eventSet.Last().AbsoluteTime;
                var deltaTime = time - lastTime;
                if (deltaTime != 0)
                {
                    float sleepTime = deltaTime * 60000f / (currentTempo * songDeltaTicks);
                    lines.Add($"\tSleepSleep({sleepTime})");
                }
                var noteEvents = eventSet.OfType<NoteEvent>()
                                         .OrderBy(noteEvent => noteEvent.NoteName[0])
                                         .OrderBy(noteEvent => noteEvent.NoteName[1] == '#')
                                         .OrderBy(noteEvent => noteEvent.CommandCode == MidiCommandCode.NoteOn);

                //Console.WriteLine($"Events @ {time}, {orderedEvents.IndexOf(eventSet)}, {noteEvents.Count()}");
                //foreach (var note in noteEvents)
                //{
                //    var upOrDown = (note.CommandCode == MidiCommandCode.NoteOn ? "Down" : "Up");
                //    Console.WriteLine($"\t{note.NoteName} {upOrDown}, {GetNoteData(note)}");
                //}

                string upLine = "\t\tSendInput, ";
                if (lines[lines.Count - 2].Contains("{vkA0sc02A Down}"))
                    upLine += "{vkA0sc02A Up}";
                upLine += GetSendString(true, noteEvents.Where(note => note.CommandCode == MidiCommandCode.NoteOff));
                if (upLine != "\t\tSendInput, ")
                {
                    lines.Add(upLine);
                }


                string downLine = "\t\tSendInput, ";
                downLine += GetSendString(false, noteEvents.Where(note => note.CommandCode == MidiCommandCode.NoteOn)
                                                         .Where(note => note.NoteName[1] != '#'));
                var shiftLines = noteEvents.Where(note => note.CommandCode == MidiCommandCode.NoteOn)
                                           .Where(note => note.NoteName[1] == '#');
                if (shiftLines.Any())
                {
                    downLine += "{vkA0sc02A Down}";
                    downLine += GetSendString(false, shiftLines);
                }
                if (downLine != "\t\tSendInput, ")
                    lines.Add(downLine);

                lastTime = time;
                if (eventSet.OfType<TempoEvent>().Count() > 0)
                    currentTempo = (float)eventSet.OfType<TempoEvent>().Last().Tempo;
            }
            lines.Add("}");
            lines.Add("\n\nSleepSleep(s)" +
            "{\n" +
            "\tSetBatchLines, -1\n" +
            "\tDllCall(\"QueryPerformanceFrequency\", \"Int64*\", freq)" +
            "\tDllCall(\"QueryPerformanceCounter\", \"Int64*\", CounterBefore)\n   " +
            "\tWhile (((counterAfter - CounterBefore) / freq * 1000) < s)\n" +
            "\tDllCall(\"QueryPerformanceCounter\", \"Int64*\", CounterAfter)\n\n" +
            "\treturn ((counterAfter - CounterBefore) / freq * 1000)\n" +
            "}");
            File.WriteAllLines(@directory + "\\" + midiNames[songCount] + ".ahk", lines);
            Console.Read();
        }

        private static string GetSendString(bool up, IEnumerable<NoteEvent> noteEvents)
        {
            string upOrDown = up ? "Up" : "Down";
            return noteEvents.Select(note => GetNoteData(note))
                             .Where(nd => nd != "")
                             .Distinct()
                             .Select(nd => "{" + $"{nd} {upOrDown}" + "}")
                             .DefaultIfEmpty()
                             .Aggregate((current, next) => current + next);
        }

        private static string GetNoteData(NoteEvent note)
        {

            //note.NoteNumber is 0 (C-1) to 120 (C9)
            int noteNumber = note.NoteNumber;
            noteNumber += transpose * 12 - 24;
            if (noteNumber < 0 || noteNumber > 60)
                return "";

            int pos = noteNumber / 12 * 7; // starts with the octave

            //this gets the note without the flat number
            switch(note.NoteName[0])
            {
                case 'D':
                    pos += 1;
                    break;
                case 'E':
                    pos += 2;
                    break;
                case 'F':
                    pos += 3;
                    break;
                case 'G':
                    pos += 4;
                    break;
                case 'A':
                    pos += 5;
                    break;
                case 'B':
                    pos += 6;
                    break;
            }


            string v = note.NoteName[0].ToString() + (IsFlat(noteNumber) ? "#" : "") + ((noteNumber + 24) / 12);
            if(note.CommandCode == MidiCommandCode.NoteOff)
            Console.WriteLine($"Note is {v}, pos is {pos}, playing {keyValues[pos]}");

            return keyValues[pos];
        }

        private static bool IsFlat(int noteNumber)
        {
            switch (noteNumber % 12)
            {
                case 1: goto IsFlat;
                case 3: goto IsFlat;
                case 6: goto IsFlat;
                case 8: goto IsFlat;
                case 10: goto IsFlat;
            }
            return false;
        IsFlat:
            return true;
        }
    }
}