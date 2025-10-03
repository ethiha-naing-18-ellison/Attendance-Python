# 🕐 5-Minute Rule - Duplicate Punch Filter

## Problem Statement

Employees sometimes accidentally punch the attendance system multiple times within a very short period (1-3 minutes), creating false duplicate entries.

### Example of the Problem:

```
Employee: THIHA NAING (ID: 263)
Punch 1: 08:30 (Check In)
Punch 2: 08:31 (Check Out) ❌ Only 1 minute later!
Punch 3: 08:32 (Check In)
Punch 4: 18:00 (Check Out)
```

This is clearly wrong - the employee didn't actually work for only 1 minute!

## Solution: 5-Minute Rule

The system now automatically detects and filters out these duplicate punches.

### The Rule:

**If the first "In" punch and first "Out" punch are less than 5 minutes (300 seconds) apart, the system will skip that "Out" punch and shift all subsequent punches forward.**

## How It Works

### Example 1: Filtering Duplicate Punch

**Raw Database Punches:**
```
Punch 1: 08:30 (In)
Punch 2: 08:32 (Out)   ← Only 2 minutes! Duplicate!
Punch 3: 12:00 (In)
Punch 4: 13:00 (Out)
Punch 5: 13:30 (In)
Punch 6: 18:00 (Out)
```

**After 5-Minute Rule Applied:**
```
In:  08:30
Out: 12:00  ← Shifted from Punch 3
In:  13:00  ← Shifted from Punch 4
Out: 13:30  ← Shifted from Punch 5
In:  18:00  ← Shifted from Punch 6
Out: (empty)
```

The **08:32** punch is skipped because it's only 2 minutes after the first punch!

### Example 2: Normal Case (No Filtering)

**Raw Database Punches:**
```
Punch 1: 08:30 (In)
Punch 2: 12:00 (Out)   ← 3.5 hours later, OK!
Punch 3: 13:00 (In)
Punch 4: 18:00 (Out)
```

**After 5-Minute Rule Applied:**
```
In:  08:30
Out: 12:00  ← No change, gap is > 5 minutes
In:  13:00
Out: 18:00
```

No filtering needed - all punches are valid!

## Excel Output

### Before (With Duplicates):

| Employee ID | Name        | In    | Out   | In    | Out   | In    | Out   |
|-------------|-------------|-------|-------|-------|-------|-------|-------|
| 263         | THIHA NAING | 08:30 | 08:31 | 08:32 | 18:00 |       |       |

❌ **Wrong!** Shows employee worked only 1 minute in the morning.

### After (With 5-Minute Rule):

| Employee ID | Name        | In    | Out   | In    | Out   | In    | Out   |
|-------------|-------------|-------|-------|-------|-------|-------|-------|
| 263         | THIHA NAING | 08:30 | 18:00 |       |       |       |       |

✅ **Correct!** Shows actual work pattern (duplicate 08:31 and 08:32 removed).

## Technical Details

### SQL Implementation

The system uses a SQL Common Table Expression (CTE) to:

1. **Fetch all punches** ordered by time
2. **Calculate time gap** between first and second punch
3. **Check if gap < 300 seconds** (5 minutes)
4. **If yes**: Skip the second punch and shift all subsequent punches forward
5. **If no**: Keep all punches as-is

### Time Calculation

```sql
strftime('%s', full_punch_2) - strftime('%s', full_punch_1) < 300
```

- Converts times to seconds since epoch
- Calculates difference
- Checks if less than 300 seconds (5 minutes)

## Why 5 Minutes?

| Time Gap | Meaning | Action |
|----------|---------|--------|
| < 1 min  | Definitely duplicate | ❌ Filter out |
| 1-2 min  | Likely duplicate | ❌ Filter out |
| 2-5 min  | Possibly duplicate | ❌ Filter out (safe) |
| 5+ min   | Legitimate punch | ✅ Keep |

**5 minutes is a safe threshold** because:
- Real work sessions are rarely less than 5 minutes
- Accidental double-punches usually happen within 1-3 minutes
- Provides buffer for edge cases

## Scope of the Rule

### What IS Filtered:
- ✅ **First In → First Out** gap only
- ✅ Only when gap < 5 minutes
- ✅ Automatically shifts remaining punches

### What IS NOT Filtered:
- ❌ Second In → Second Out (not checked)
- ❌ Third In → Third Out (not checked)
- ❌ Gaps > 5 minutes (kept as-is)

## Real-World Examples

### Case 1: Double Punch at Morning Clock-In

**Raw:**
```
08:30 (In)
08:31 (Out) ← Accidentally punched again
08:32 (In)  ← Realized mistake, punched correctly
12:00 (Out)
13:00 (In)
18:00 (Out)
```

**Filtered:**
```
In:  08:30
Out: 08:32 (shifted from punch 3)
In:  12:00 (shifted from punch 4)
Out: 13:00 (shifted from punch 5)
In:  18:00 (shifted from punch 6)
```

### Case 2: Quick Test Punch

**Raw:**
```
08:28 (In)
08:29 (Out) ← Testing if device works
08:30 (In)  ← Actual clock-in
18:00 (Out)
```

**Filtered:**
```
In:  08:28
Out: 08:30 (shifted)
In:  18:00 (shifted)
Out: (empty)
```

## Benefits

1. ✅ **Cleaner Data**: Removes obvious errors
2. ✅ **Accurate Reports**: Shows real work patterns
3. ✅ **Automatic**: No manual intervention needed
4. ✅ **Transparent**: Logic is clear and documented
5. ✅ **Safe**: Conservative 5-minute threshold

## Limitations

### What This Does NOT Handle:

1. **Multiple short duplicates**: Only filters first duplicate
2. **Legitimate short sessions**: If someone actually worked 3 minutes, it will be filtered
3. **Missing punches**: Doesn't add missing data
4. **Other In/Out pairs**: Only checks first pair

### Recommendation:

- For complex attendance issues, use the full attendance system
- This tool is for **raw data export** with basic filtering
- Manual review may still be needed for edge cases

## Configuration

### Current Settings:

```python
TIME_THRESHOLD = 300  # 5 minutes = 300 seconds
```

To change the threshold, modify the SQL query in `data_generator_api.py`:

```sql
AND (strftime('%s', full_punch_2) - strftime('%s', full_punch_1)) < 300
                                                                    ^^^
                                                            Change this value
```

**Common thresholds:**
- 180 = 3 minutes (stricter)
- 300 = 5 minutes (default, balanced)
- 600 = 10 minutes (more lenient)

## Testing

### Test Case 1: Duplicate Within 2 Minutes

Input:
```
263, THIHA NAING, 08:30, 08:32
```

Expected Output:
```
263, THIHA NAING, 08:30, (shifted to next punch)
```

### Test Case 2: Normal 6-Hour Gap

Input:
```
263, THIHA NAING, 08:30, 14:30
```

Expected Output:
```
263, THIHA NAING, 08:30, 14:30 (no change)
```

## Troubleshooting

### Q: My legitimate 4-minute session was filtered!

**A:** The 5-minute rule is designed to catch duplicates. If you have many legitimate short sessions, consider:
- Using the main attendance API (more flexible)
- Adjusting the threshold to 2-3 minutes
- Manually reviewing filtered data

### Q: I still see duplicates in my report!

**A:** The rule only checks the **first In/Out pair**. If duplicates appear later in the day, they won't be filtered. This is by design to avoid over-filtering.

### Q: How can I see what was filtered?

**A:** Currently, filtered punches are silently skipped. For debugging, you could:
- Check the raw database before filtering
- Compare with the Excel output
- Add logging to the API (developer option)

## Summary

✅ **5-Minute Rule = Smart Duplicate Filter**

- Automatically removes obvious duplicate punches
- Applies only to first In/Out pair
- Shifts remaining punches forward
- Makes data cleaner and more accurate
- Safe and conservative threshold

---

**Version**: 1.0.0  
**Implemented**: October 2025  
**Threshold**: 5 minutes (300 seconds)

