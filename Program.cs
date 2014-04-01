using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace ConsoleApp
{
	class Program
	{
		static void Main(string[] args)
		{
			var pyconTimezone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");

			var doc = GetHtmlDocument(new Uri("https://us.pycon.org/2014/schedule/talks/list/"));

			var formattedTalks = GetTalks(pyconTimezone, doc);
			Console.WriteLine("There are {0} talks, writing to iCal file ... ", formattedTalks.Count());

			WriteICal(pyconTimezone, formattedTalks, @"pycon2014.ics");
			Console.WriteLine("Done!");

			Console.ReadLine();
		}

		private static void WriteICal(TimeZoneInfo timezone, Talk[] formattedTalks, string icalFilename)
		{
			if (File.Exists(icalFilename))
				File.Delete(icalFilename);
			var headerLines = GetICalHeaderLines(timezone).ToArray();
			var footerLines = GetICalFooterLines().ToArray();

			using (var streamWriter = new StreamWriter(File.OpenWrite(icalFilename)))
			{
				WriteLines(headerLines, streamWriter);

				foreach (var talk in formattedTalks)
				{
					WriteLines(GetICalLines(talk, timezone), streamWriter);
				}

				WriteLines(footerLines, streamWriter);
			}
		}

		private static Talk[] GetTalks(TimeZoneInfo timezone, HtmlDocument doc)
		{
			var talks = doc.DocumentNode.SelectNodes("//div[@class='span8 presentation well']");

			var formattedTalks = (from talk in talks
								  select ConvertTalk(talk, timezone)).ToArray();
			return formattedTalks;
		}

		private static HtmlDocument GetHtmlDocument(Uri url)
		{
			HtmlDocument doc = new HtmlDocument();
			string text;
			using (var client = new WebClient())
			{
				client.Encoding = Encoding.UTF8;
				text = client.DownloadString(url);
			}
			doc.LoadHtml(text);
			return doc;
		}

		private static void WriteLines(string[] totalLines, StreamWriter streamWriter)
		{
			foreach (var line in totalLines)
				streamWriter.WriteLine(line);
		}

		private static IEnumerable<string> GetICalFooterLines()
		{
			yield return "END:VCALENDAR";
		}

		private static IEnumerable<string> GetICalHeaderLines(TimeZoneInfo timezone)
		{
			yield return "BEGIN:VCALENDAR";
			yield return "PRODID:-//djeebus/pycon2014//NONSGML v1.0//EN";
			yield return "VERSION:2.0";
			yield return "CALSCALE:GREGORIAN";
			yield return "X-WR-CALNAME:Pycon 2014 Schedule";
			yield return string.Format("X-WR-TIMEZONE:{0}", FormatTimeZone(timezone));
			//yield return "BEGIN:VTIMEZONE";
			//yield return string.Format("TZID:{0}", FormatTimeZone(timezone));

			//foreach (var rule in timezone.GetAdjustmentRules())
			//{
			//	if (timezone.SupportsDaylightSavingTime)
			//	{
			//		yield return "BEGIN:DAYLIGHT";
			//		yield return string.Format("DTSTART:{0}", FormatDateTime(rule.DateStart, timezone));
			//		yield return string.Format("RRULE:{0}", FormatRule(timezone, rule, rule.DaylightTransitionStart));
			//		yield return string.Format("TZOFFSETFROM:{0}", FormatTimeSpan(timezone.BaseUtcOffset));
			//		yield return string.Format("TZOFFSETTO:{0}", FormatTimeSpan(timezone.BaseUtcOffset + rule.DaylightDelta));
			//		yield return "END:DAYLIGHT";
			//	}

			//	yield return "BEGIN:STANDARD";
			//	yield return string.Format("DTSTART:{0}", FormatDateTime(rule.DateStart, timezone));
			//	yield return string.Format("RRULE:{0}", FormatRule(timezone, rule, rule.DaylightTransitionEnd));
			//	yield return string.Format("TZOFFSETFROM:{0}", FormatTimeSpan(timezone.BaseUtcOffset + rule.DaylightDelta));
			//	yield return string.Format("TZOFFSETTO:{0}", FormatTimeSpan(timezone.BaseUtcOffset));
			//	yield return "END:STANDARD";
			//}

			//yield return "END:VTIMEZONE";

			// hard coded, but the above didn't work *sigh*
			yield return string.Format(@"BEGIN:VTIMEZONE
TZID:{0}
BEGIN:STANDARD
DTSTART:16011104T020000
RRULE:FREQ=YEARLY;BYDAY=1SU;BYMONTH=11
TZOFFSETFROM:-0400
TZOFFSETTO:-0500
END:STANDARD
BEGIN:DAYLIGHT
DTSTART:16010311T020000
RRULE:FREQ=YEARLY;BYDAY=2SU;BYMONTH=3
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
END:DAYLIGHT
END:VTIMEZONE", FormatTimeZone(timezone));
		}

		private static string FormatTimeSpan(TimeSpan timespan)
		{
			return string.Format("{0:D2}{1:D2}", timespan.Hours, timespan.Minutes);
		}

		private static string FormatRule(TimeZoneInfo timezone, TimeZoneInfo.AdjustmentRule rule, TimeZoneInfo.TransitionTime trans)
		{
			return string.Format("FREQ=YEARLY;BYMONTH={0};BYDAY={1}{2};UNTIL={3}",
				trans.Month,
				trans.Week,
				trans.DayOfWeek.ToString().Substring(0, 2).ToUpper(),
				FormatDateTime(rule.DateEnd, timezone));
		}

		static readonly DateTime MinDate = new DateTime(1700, 1, 1);
		private static string FormatDateTime(DateTime datetime, TimeZoneInfo timeZone)
		{
			if (datetime < MinDate)
				datetime = MinDate;

			var isUtc = true;
			if (datetime.Kind != DateTimeKind.Unspecified)
			{
				datetime = TimeZoneInfo.ConvertTimeFromUtc(datetime, timeZone);
				isUtc = false;
			}

			return string.Format("{0:yyyyMMdd}T{0:HHmmss}{1}", datetime, isUtc ? "Z" : "");
		}

		private static string FormatTimeZone(TimeZoneInfo timezone)
		{
			return timezone.Id; // required for outlook importing
		}

		private static string[] GetICalLines(Talk talk, TimeZoneInfo timezone)
		{
			return new[] 
			{
				"BEGIN:VEVENT",
				string.Format("DTSTART;TZID=\"{1}\":{0}", FormatDateTime(talk.Start, timezone), FormatTimeZone(timezone)),
				string.Format("DTEND;TZID=\"{1}\":{0}", FormatDateTime(talk.End, timezone), FormatTimeZone(timezone)),
				string.Format("DTSTAMP:{0}", FormatDateTime(DateTime.UtcNow, timezone)),
				string.Format("UID:pycon-2014-2.{0}", talk.Id),
				string.Format("CREATED:{0}", FormatDateTime(DateTime.UtcNow, timezone)),
				string.Format("DESCRIPTION:{0}\n\n{1}", (talk.Description ?? string.Empty).Replace(",", @"\,"), talk.Url),
				string.Format("LOCATION:{0}", talk.Location),
				string.Format("X-ALT-DESC;FMTTYPE=text/html:<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 3.2//EN\"><HTML><BODY>{0}<br><a href='{1}'>{1}</a></BODY></HTML>", talk.Description.Replace(",", @"\,"), talk.Url),
				string.Format("LAST-MODIFIED:{0}", FormatDateTime(DateTime.UtcNow, timezone)),
				string.Format("URL:{0}", talk.Url),
				string.Format("SUMMARY:{0}", talk.Title),
				"SEQUENCE:1",

				"STATUS:CONFIRMED",
				"X-MICROSOFT-CDO-BUSYSTATUS:FREE",
				"TRANSP:TRANSPARENT",
				
				"END:VEVENT",
			};
		}

		class Talk
		{
			public int Id { get; set; }
			public static readonly Uri RootUri = new Uri("https://us.pycon.org/2014/schedule/talks/list/");
			public string Title { get; set; }
			public Uri Url { get; set; }
			public string Subtitle { get; set; }
			public string Description { get; set; }
			public DateTime Start { get; set; }
			public DateTime End { get; set; }
			public string Location { get; set; }
		}

		static readonly Dictionary<string, DateTime> _dateMap = new Dictionary<string, DateTime>
		{
			{ "Friday", new DateTime(2014, 4, 11) },
			{ "Saturday", new DateTime(2014, 4, 12) },
			{ "Sunday", new DateTime(2014, 4, 13) },
		};

		private static Talk ConvertTalk(HtmlNode node, TimeZoneInfo timezone)
		{
			var talk = new Talk();

			var titleNode = node.SelectSingleNode("h3");
			var link = titleNode.SelectSingleNode("a").Attributes["href"].Value;
			var title = titleNode.InnerText;

			var subtitleNode = GetNextNode(titleNode);
			var subtitle = subtitleNode.InnerText;

			var descriptionNode = GetNextNode(subtitleNode);
			var description = descriptionNode.InnerText.Trim().Replace("\r\n", " ");

			var locationAndTimesNode = GetNextNode(descriptionNode);
			var locationAndTimes = locationAndTimesNode.InnerText;

			var parts = locationAndTimes.Trim().Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
			parts = (from p in parts
					 select p.Trim()).ToArray();
			var timeParts = parts[1].Split(new[] { "&ndash;" }, StringSplitOptions.None);

			var day = _dateMap[parts[0]];

			talk.Title = CleanEscapeCharacters(title);
			talk.Url = new Uri(Talk.RootUri, link);
			talk.Description = CleanEscapeCharacters(description);
			talk.Subtitle = CleanEscapeCharacters(subtitle);
			talk.Start = ParseTime(day, timeParts[0], timezone);
			talk.End = ParseTime(day, timeParts[1], timezone);
			talk.Location = CleanEscapeCharacters(parts[3]);

			talk.Id = int.Parse(talk.Url.Segments[4].Replace("/", string.Empty));

			return talk;
		}


		//static readonly Regex _hexadecimalNumericCharacterReference = new Regex(
		//	@"\&\#([\dA-Fa-f]+)\;", RegexOptions.Compiled | RegexOptions.Multiline);
		private static string CleanEscapeCharacters(string value)
		{
			//var hexadecimalProcessor = new MatchEvaluator(m =>
			//	char.ConvertFromUtf32(int.Parse(m.Groups[1].Value, NumberStyles.HexNumber)));
			//value = _hexadecimalNumericCharacterReference.Replace(value, hexadecimalProcessor);
			return value
				.Replace("&#39;", "'")
				.Replace("â€™", "'")
				.Replace("&amp;", "&")
				.Replace("&quot;", "'");
		}

		private static DateTime ParseTime(DateTime day, string time, TimeZoneInfo timezone)
		{
			time = time
				.Replace("a.m.", "AM")
				.Replace("p.m.", "PM")
				.Replace("noon", "12:00 PM");
			if (!time.Contains(':'))
				time = time.Replace(" ", ":00 ");
			time = string.Format("{0:yyyy} {0:MM} {0:dd} {1}", day, time);
			
			var dt = DateTime.ParseExact(time, "yyyy MM dd h:mm tt", CultureInfo.InvariantCulture);
			var local = TimeZoneInfo.ConvertTimeFromUtc(dt, TimeZoneInfo.Utc);
			var utc = TimeZoneInfo.ConvertTimeToUtc(dt, timezone);
			return utc;
		}

		static readonly Regex _whitespace = new Regex(@"\s+", RegexOptions.Compiled);
		private static HtmlNode GetNextNode(HtmlNode titleNode)
		{
			var node = titleNode;

			while (true)
			{
				node = node.NextSibling;

				if (node == null)
					return null;

				switch (node.NodeType)
				{
					case HtmlNodeType.Comment:
						continue;

					case HtmlNodeType.Document:
						throw new ArgumentException("titleNode", "How'd we get a document from this??");

					case HtmlNodeType.Element:
						return node;

					case HtmlNodeType.Text:
						var text = _whitespace.Replace(node.InnerText, "");
						if (string.IsNullOrEmpty(text))
							continue;

						return node;
				}
			}
		}
	}
}
