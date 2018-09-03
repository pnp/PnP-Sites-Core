using System;

namespace OfficeDevPnP.Core.Enums
{
    /// <summary>
    /// Timezones to use when creating sitecollections
    /// Format UTC[PLUS|MINUS][HH:MM]_[DESCRIPTION]
    /// </summary>
    public enum TimeZone
    {
        /// <summary>
        /// No Timezone
        /// </summary>
        None = 0,
        /// <summary>
        /// Timezone for GREENWICH, DUBLIN, EDINBURGH, LISBON and LONDON
        /// </summary>
        UTC_GREENWICH_MEAN_TIME_DUBLIN_EDINBURGH_LISBON_LONDON = 2,
        /// <summary>
        /// Timezone for BRUSSELS, COPENHAGEN, MADRID and PARIS
        /// </summary>
        UTCPLUS0100_BRUSSELS_COPENHAGEN_MADRID_PARIS = 3,
        /// <summary>
        /// Timezone for AMSTERDAM, BERLIN, BERN, ROME, STOCKHOLM and VIENNA
        /// </summary>
        UTCPLUS0100_AMSTERDAM_BERLIN_BERN_ROME_STOCKHOLM_VIENNA = 4,
        /// <summary>
        /// Timezone for ATHENS, BUCHAREST and ISTANBUL
        /// </summary>
        UTCPLUS0200_ATHENS_BUCHAREST_ISTANBUL = 5,
        /// <summary>
        /// Timezone for BELGRADE, BRATISLAVA, BUDAPEST, LJUBLJANA and PRAGUE
        /// </summary>
        UTCPLUS0100_BELGRADE_BRATISLAVA_BUDAPEST_LJUBLJANA_PRAGUE = 6,
        /// <summary>
        /// Timezone for MINSK 
        /// </summary>
        UTCPLUS0200_MINSK = 7,
        /// <summary>
        /// Timezone for BRASILIA
        /// </summary>
        UTCMINUS0300_BRASILIA = 8,
        /// <summary>
        /// Timezone for ATLANTIC CANADA
        /// </summary>
        UTCMINUS0400_ATLANTIC_TIME_CANADA = 9,
        /// <summary>
        /// Timezone for EASTERN US and CANADA
        /// </summary>
        UTCMINUS0500_EASTERN_TIME_US_AND_CANADA = 10,
        /// <summary>
        /// Timezone for CENTRAL US and CANADA
        /// </summary>
        UTCMINUS0600_CENTRAL_TIME_US_AND_CANADA = 11,
        /// <summary>
        /// Timezone for MOUNTAIN US and CANADA
        /// </summary>
        UTCMINUS0700_MOUNTAIN_TIME_US_AND_CANADA = 12,
        /// <summary>
        /// Timezone for PACIFIC US and CANADA
        /// </summary>
        UTCMINUS0800_PACIFIC_TIME_US_AND_CANADA = 13,
        /// <summary>
        /// Timezone for ALASKA
        /// </summary>
        UTCMINUS0900_ALASKA = 14,
        /// <summary>
        /// Timezone for HAWAII
        /// </summary>
        UTCMINUS1000_HAWAII = 15,
        /// <summary>
        /// Timezone for MIDWAY ISLAND and SAMOA
        /// </summary>
        UTCMINUS1100_MIDWAY_ISLAND_SAMOA = 16,
        /// <summary>
        /// Timezone for AUCKLAND and WELLINGTON
        /// </summary>
        [Obsolete("Use UTCPLUS1200_AUCKLAND_WELLINGTON instead")]
        UTCPLUS1200_AUKLAND_WELLINGTON = 17,
        /// <summary>
        /// Timezone for AUCKLAND and WELLINGTON
        /// </summary>
        UTCPLUS1200_AUCKLAND_WELLINGTON = 17,
        /// <summary>
        /// Timezone for BRISBANE
        /// </summary>
        UTCPLUS1000_BRISBANE = 18,
        /// <summary>
        /// Timezone for ADELAIDE
        /// </summary>
        UTCPLUS0930_ADELAIDE = 19,
        /// <summary>
        /// Timezone for OSAKA, SAPPORO and TOKYO
        /// </summary>
        UTCPLUS0900_OSAKA_SAPPORO_TOKYO = 20,
        /// <summary>
        /// Timezone for KUALA LUMPUR and SINGAPORE
        /// </summary>
        UTCPLUS0800_KUALA_LUMPUR_SINGAPORE = 21,
        /// <summary>
        /// Timezone for BANGKOK, HANOI and JAKARTA
        /// </summary>
        UTCPLUS0700_BANGKOK_HANOI_JAKARTA = 22,
        /// <summary>
        /// Timezone for CHENNAI, KOLKATA, MUMBAI and NEW DELHI
        /// </summary>
        UTCPLUS0530_CHENNAI_KOLKATA_MUMBAI_NEW_DELHI = 23,
        /// <summary>
        /// Timezone for ABU DHABI and MUSCAT
        /// </summary>
        UTCPLUS0400_ABU_DHABI_MUSCAT = 24,
        /// <summary>
        /// Timezone for TEHRAN
        /// </summary>
        UTCPLUS0330_TEHRAN = 25,
        /// <summary>
        /// Timezone for BAGHDAD
        /// </summary>
        UTCPLUS0300_BAGHDAD = 26,
        /// <summary>
        /// Timezone for JERUSALEM
        /// </summary>
        UTCPLUS0200_JERUSALEM = 27,
        /// <summary>
        /// Timezone for NEWFOUNDLAND and LABRADOR
        /// </summary>
        UTCMINUS0330_NEWFOUNDLAND_AND_LABRADOR = 28,
        /// <summary>
        /// Timezone for AZORES
        /// </summary>
        UTCMINUS0100_AZORES = 29,
        /// <summary>
        /// Timezone for MIDATLANTIC
        /// </summary>
        UTCMINUS0200_MID_ATLANTIC = 30,
        /// <summary>
        /// Timezone for MONROVIA
        /// </summary>
        UTC_MONROVIA = 31,
        /// <summary>
        /// Timezone for CAYENNE
        /// </summary>
        UTCMINUS0300_CAYENNE = 32,
        /// <summary>
        /// Timezone for GEORGETOWN and LA PAZ - SAN JUAN
        /// </summary>
        UTCMINUS0400_GEORGETOWN_LA_PAZ_SAN_JUAN = 33,
        /// <summary>
        /// Timezone for INDIANA EAST
        /// </summary>
        UTCMINUS0500_INDIANA_EAST = 34,
        /// <summary>
        /// Timezone for BOGOTA, LIMA and QUITO
        /// </summary>
        UTCMINUS0500_BOGOTA_LIMA_QUITO = 35,
        /// <summary>
        /// Timezone for SASKATCHEWAN
        /// </summary>
        UTCMINUS0600_SASKATCHEWAN = 36,
        /// <summary>
        /// Timezone for GUADALAJARA, MEXICO CITY and MONTERREY
        /// </summary>
        UTCMINUS0600_GUADALAJARA_MEXICO_CITY_MONTERREY = 37,
        /// <summary>
        /// Timezone for ARIZONA
        /// </summary>
        UTCMINUS0700_ARIZONA = 38,
        /// <summary>
        /// Timezone for INTERNATIONAL DATE LINE WEST
        /// </summary>
        UTCMINUS1200_INTERNATIONAL_DATE_LINE_WEST = 39,
        /// <summary>
        /// Timezone for FIJI ISLANDS and MARSHALL ISLAND
        /// </summary>
        UTCPLUS1200_FIJI_ISLANDS_MARSHALL_ISLANDS = 40,
        /// <summary>
        /// Timezone for MADAGAN, SOLOMON ISLANDS and NEW CALENDONIA
        /// </summary>
        UTCPLUS1100_MADAGAN_SOLOMON_ISLANDS_NEW_CALENDONIA = 41,
        /// <summary>
        /// Timezone for HOBART
        /// </summary>
        UTCPLUS1000_HOBART = 42,
        /// <summary>
        /// Timezone for GUAM and PORT MORESBY
        /// </summary>
        UTCPLUS1000_GUAM_PORT_MORESBY = 43,
        /// <summary>
        /// Timezone for DARWIN
        /// </summary>
        UTCPLUS0930_DARWIN = 44,
        /// <summary>
        /// Timezone for BEIJING, CHONGQING, HONG KONG and SAR URUMQI
        /// </summary>
        UTCPLUS0800_BEIJING_CHONGQING_HONG_KONG_SAR_URUMQI = 45,
        /// <summary>
        /// Timezone for NOVOSIBIRSK
        /// </summary>
        UTCPLUS0600_NOVOSIBIRSK = 46,
        /// <summary>
        /// Timezone for TASHKENT
        /// </summary>
        UTCPLUS0500_TASHKENT = 47,
        /// <summary>
        /// Timezone for KABUL
        /// </summary>
        UTCPLUS0430_KABUL = 48,
        /// <summary>
        /// Timezone for CAIRO
        /// </summary>
        UTCPLUS0200_CAIRO = 49,
        /// <summary>
        /// Timezone for HARARE and PRETORIA
        /// </summary>
        UTCPLUS0200_HARARE_PRETORIA = 50,
        /// <summary>
        /// Timezone for MOSCOW, ST PETERSBURG and VOLGOGRAD
        /// </summary>
        UTCPLUS0300_MOSCOW_STPETERSBURG_VOLGOGRAD = 51,
        /// <summary>
        /// Timezone for CAPE VERDE ISLANDS
        /// </summary>
        UTCMINUS0100_CAPE_VERDE_ISLANDS = 53,
        /// <summary>
        /// Timezone for BAKU
        /// </summary>
        UTCPLUS0400_BAKU = 54,
        /// <summary>
        /// Timezone for CENTAL AMERICA
        /// </summary>
        UTCMINUS0600_CENTRAL_AMERICA = 55,
        /// <summary>
        /// Timezone for NAIROBI
        /// </summary>
        UTCPLUS0300_NAIROBI = 56,
        /// <summary>
        /// Timezone for SARAJEVO, SKOPJE, WARSAW and ZAGREB
        /// </summary>
        UTCPLUS0100_SARAJEVO_SKOPJE_WARSAW_ZAGREB = 57,
        /// <summary>
        /// Timezone for EKATERINBURG
        /// </summary>
        UTCPLUS0500_EKATERINBURG = 58,
        /// <summary>
        /// Timezone for HELSINKI, KYIV, RIGA, SOFIA, TALLINN and VILNIUS
        /// </summary>
        UTCPLUS0200_HELSINKI_KYIV_RIGA_SOFIA_TALLINN_VILNIUS = 59,
        /// <summary>
        /// Timezone for GREENLAND
        /// </summary>
        UTCMINUS0300_GREENLAND = 60,
        /// <summary>
        /// Timezone for YANGON and RANGOON
        /// </summary>
        UTCPLUS0630_YANGON_RANGOON = 61,
        /// <summary>
        /// Timezone for KATHMANDU
        /// </summary>
        UTCPLUS0545_KATHMANDU = 62,
        /// <summary>
        /// Timezone for IRKUTSK
        /// </summary>
        UTCPLUS0800_IRKUTSK = 63,
        /// <summary>
        /// Timezone for KRASNOYARSK
        /// </summary>
        UTCPLUS0700_KRASNOYARSK = 64,
        /// <summary>
        /// Timezone for SANTIAGO
        /// </summary>
        UTCMINUS0400_SANTIAGO = 65,
        /// <summary>
        /// Timezone for SRI JAYAWARDENEPURA
        /// </summary>
        UTCPLUS0530_SRI_JAYAWARDENEPURA = 66,
        /// <summary>
        /// Timezone for NUKU and ALOFA
        /// </summary>
        UTCPLUS1300_NUKU_ALOFA = 67,
        /// <summary>
        /// Timezone for VLADIVOSTOK
        /// </summary>
        UTCPLUS1000_VLADIVOSTOK = 68,
        /// <summary>
        /// Timezone for WEST CENTRAL AFRICA 
        /// </summary>
        UTCPLUS0100_WEST_CENTRAL_AFRICA = 69,
        /// <summary>
        /// Timezone for YAKUTSK
        /// </summary>
        UTCPLUS0900_YAKUTSK = 70,
        /// <summary>
        /// Timezone for ASTANA and DHAKA
        /// </summary>
        UTCPLUS0600_ASTANA_DHAKA = 71,
        /// <summary>
        /// Timezone for SEOUL 
        /// </summary>
        UTCPLUS0900_SEOUL = 72,
        /// <summary>
        /// Timezone for PERTH
        /// </summary>
        UTCPLUS0800_PERTH = 73,
        /// <summary>
        /// Timezone for KUWAIT and RIYADH
        /// </summary>
        UTCPLUS0300_KUWAIT_RIYADH = 74,
        /// <summary>
        /// Timezone for TAIPEI
        /// </summary>
        UTCPLUS0800_TAIPEI = 75,
        /// <summary>
        /// Timezone for CANBERRA, MELBOURNE and SYDNEY
        /// </summary>
        UTCPLUS1000_CANBERRA_MELBOURNE_SYDNEY = 76,
        /// <summary>
        /// Timezone for CHIHUAHUA and LA PAZ - MAZATLAN
        /// </summary>
        UTCMINUS0700_CHIHUAHUA_LA_PAZ_MAZATLAN = 77,
        /// <summary>
        /// Timezone for TIJUANA, BAJA and CALFORNIA
        /// </summary>
        UTCMINUS0800_TIJUANA_BAJA_CALFORNIA = 78,
        /// <summary>
        /// Timezone for AMMAN
        /// </summary>
        UTCPLUS0200_AMMAN = 79,
        /// <summary>
        /// Timezone for BEIRUT
        /// </summary>
        UTCPLUS0200_BEIRUT = 80,
        /// <summary>
        /// Timezone for MANAUS
        /// </summary>
        UTCMINUS0400_MANAUS = 81,
        /// <summary>
        /// Timezone for TBILISI
        /// </summary>
        UTCPLUS0400_TBILISI = 82,
        /// <summary>
        /// Timezone for WINDHOEK
        /// </summary>
        UTCPLUS0200_WINDHOEK = 83,
        /// <summary>
        /// Timezone for YEREVAN
        /// </summary>
        UTCPLUS0400_YEREVAN = 84,
        /// <summary>
        /// Timezone for BUENOS AIRES
        /// </summary>
        UTCMINUS0300_BUENOS_AIRES = 85,
        /// <summary>
        /// Timezone for CASABLANCA
        /// </summary>
        UTC_CASABLANCA = 86,
        /// <summary>
        /// Timezone for ISLAMABAD and KARACHI
        /// </summary>
        UTCPLUS0500_ISLAMABAD_KARACHI = 87,
        /// <summary>
        /// Timezone for CARACAS
        /// </summary>
        UTCMINUS0430_CARACAS = 88,
        /// <summary>
        /// Timezone for PORT LOUIS
        /// </summary>
        UTCPLUS0400_PORT_LOUIS = 89,
        /// <summary>
        /// Timezone for MONTEVIDEO
        /// </summary>
        UTCMINUS0300_MONTEVIDEO = 90,
        /// <summary>
        /// Timezone for ASUNCION
        /// </summary>
        UTCMINUS0400_ASUNCION = 91,
        /// <summary>
        /// Timezone for PETROPAVLOVSK and KACHATSKY
        /// </summary>
        UTCPLUS1200_PETROPAVLOVSK_KACHATSKY = 92,
        /// <summary>
        /// COORDINATED UNIVERSAL TIME
        /// </summary>
        UTC_COORDINATED_UNIVERSAL_TIME = 93,
        /// <summary>
        /// Timezone for ULAANBAATAR
        /// </summary>
        UTCMINUS0800_ULAANBAATAR = 94
    }
}
