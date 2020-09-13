# Google Chores tool

Converts an google sheet of jobs into individual google calendar events. Useful as instead of creating 100+ individual events and inviting specific housemates to them I can just create a sheet of a particular format and run this script. Combats housemates 'forgetting' to do their chores as it can be set up to email and give notifications to them on the day they need to be done.

Takes in a sheet of the format

| Name          | Description       | Day     |07/09/2020|14/09/2020|
| ------------- |-------------------|---------|----------|--------  |
| Sweeping      | Sweep downstairs  | Monday  |HS        |EL        |
| Mopping       | Mop upstairs      | Sunday  |EL        |TS        |
| Bins          | Take bins to road | Monday  |TS HS     |HS EL     |

And will create (in this example) 6 google calendar events (one for each cell with a date above it on the right) at the date specified on row 0 plus the number of days offset specified for the job in column 2.
It will look in people.json for the specified person object in the form:

```js

{
    HS: {'email': 'example@gmail.com', name: 'Example'}
}

```