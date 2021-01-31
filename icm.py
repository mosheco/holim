import argparse

def parse_arguments():
    """
    Parse the command-line arguments.
    """
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '--payouts',
        nargs='+',
        type=int,
        help='List of payout values.',
        required=True
    )
    parser.add_argument(
        '--stacks',
        nargs='+',
        type=float,
        help='List of stacks.',
        required=True
    )

    return parser.parse_args()


def get_equity(payouts, stacks, player_index, total=None, payout_index=None):
    """ Returns the equity for a specified player based on stack size and payouts. """
    if total == None:
        total = sum(stacks)
        payout_index = 0
        stacks = list([float(stack) for stack in stacks])
        if stacks[player_index] == 0:
            #If hero's stack is 0, they come in "last place" of all
            #remaining. However if we are ITM, that doesn't = 0....!
            if len(stacks) > len(payouts):
                return 0.
            else:
                return float(payouts[len(stacks) - 1])

    ###
    # The equity for player at player_index (PL) is equal to the probabilty of winning first place multiplied by first prize +
    # the probability of player 0 winning first prize multiplied by PL's equity for winning 2nd prize or lower when player 0 has
    # won first prize +
    # the probability of player 1 winning first prize multiplied by PL's equity for winning 2nd prize when player 1 has
    # won first prize , + and so forth for all other players except PL.
    #
    # The equity for winning second prize or lower is computed recursively by removing the winning players stack .
    ###
    equity = stacks[player_index] / total * payouts[payout_index] 
    if payout_index < len(payouts) - 1: #payout_index is 0-based
        for i, stack in enumerate(stacks):
            if i != player_index and stack > 0:
                stacks[i] = 0.
                equity += get_equity(payouts, stacks, player_index, total - stack, payout_index + 1) * stack / total
                stacks[i] = stack
                
    return equity

def get_all_equities(payouts, stacks):
    """ Return list of all equities for each player defined in stacks. """
    for i in range(len(stacks)):
        yield get_equity(payouts, list(stacks), i)

if __name__ == '__main__':
    args = parse_arguments()

    for i, eq in enumerate(get_all_equities(args.payouts, args.stacks)):
        print ('Player %d: stack %d equity %.3f' % (i, args.stacks[i], eq))
