# Ideas for AnnCo 2.0

***

## Refactoring
- Compile TextGrid search patterns.
- Get rid of minidom.
- Move `ENCOD_MSG` to `Application` class.
- Make `Annotation` and `Tier` classes iterable.
- Try replacing index with re pattern in `wav_path`.
- Store Tier and Annotation elements?

## Feature
- Support for PointTiers

# 25.05.2021 Commit

***

## Support for Praat point tiers
- Tier objects now have is_point attribute of bool type defining if tier is a point tier. Default is False.
- Interval objects are considered points if their start and end attributes contain the same value.
- to_tg() method of Interval class now checks if instance is a point and returns according string representation for TextGrid
- to_tg() method of Tier class now checks if Tier instance is a point tier and and returns according string representation for TextGrid. Additionaly as `end` argument it accepts Annotation duration.
- Added point tier checkbutton to interface