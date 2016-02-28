//algorithms.cpp - Implementation of tsp class, which manages the data
//for my traveling salesman problem algorithms. Cities for the problem
//are stored as two deques of pointers to city objects, original_list
//contains the initial list of cities read from file and solution contains
//the order of the current solution to the problem.
//Written by Sam Best for CS325
//Last modified 3/12/2013
#include "stdafx.h"
#include "algorithms.h"
#include <iostream>
#include <deque>
#include <fstream>
#include <cstdlib>

using namespace std;

int done = 0;

tsp::tsp()
{
}

//Destructor clears the deques
tsp::~tsp()
{
}

void tsp::clear()
{
    for(deque <city*>::iterator it = original_list.begin(); it != original_list.end(); it++)
    {
        delete *it;
    }
    for(deque <city*>::iterator it = solution.begin(); it != solution.end(); it++)
    {
        delete *it;
    }

    original_list.clear();
    solution.clear();
}

void tsp::add(double dx, double dy)
{
    int x = (int)(dx * 1000000);
    int y = (int)(dy * 1000000);
    int n = (int)original_list.size();
    original_list.push_back(new city(n, x, y, n, false));
}

city* tsp::getSolution(int idx)
{
    return solution[(size_t)idx];
}

//Public wrapper, calls recursive brute force function
int tsp::brute_force_wrapper()
{
    int distance = 0;
    int min_distance = 9999999;
    deque <city*> best_path;

    if (solution.empty())
        nearest_neighbor_basic(0);

    brute_force(best_path, min_distance, original_list.size());
    copy_city_deque(best_path, solution); //copy best solution into solution

    distance = get_solution_distance();
    //write_solution("bruteforce.txt");
    return distance;
}

//Recursive brute force algorithm to calculate all permutations of tours and find the optimal one.
void tsp::brute_force(deque <city*> & best_path, int & min_distance, int cities_left)
{
    int current_dist = 0;

    signal(SIGTERM, end_opt);
    for (int i = 0; !done && i < cities_left; ++i)
    {
        rotate(cities_left-1);
        current_dist = get_solution_distance();
        if (current_dist < min_distance)
        {
            min_distance = current_dist;
            cout << min_distance << endl;
            copy_city_deque(solution, best_path);
        }

        brute_force(best_path, min_distance, cities_left - 1); //shift ending position of rotation left 1 in each recursion

    }

   return;
}

//Called by brute force. Moves solution[pos] to start of the solution deque
void tsp::rotate(int pos)
{
    solution.push_front(solution[pos]);
    solution.erase(solution.begin() + pos+1);

    return;
}

//Finds an initial tour using nearest neighbor, optimizes with basic 2-opt algorithm
int tsp::nearest_neighbor()
{
    solution.clear();
    int total_dist = 0;
    int best_start_distance = 9999999;
    int last_run = 0;

    //Run through entire 2-opt optimization for each city, write best index.
    for (int i = 0; !done && i < (int)original_list.size(); ++i)
    {
        nearest_neighbor_basic(i);
        last_run = two_change();
        if (last_run < best_start_distance)
        {
            best_start_distance = last_run;
//            cout << "Writing solution " << best_start_distance << endl;
//            write_solution(OUTPUT_FN);  //write each time an improvement is found
        }
    }

    total_dist = get_solution_distance();
    if (best_start_distance <= total_dist)
        return best_start_distance;     //solution already written

    else
    {
//        cout << "Writing solution " << total_dist << endl;
//        write_solution(OUTPUT_FN);       //write current solution (midway through 2-opt)
        return best_start_distance;
    }
}

//Generates nearest neighbor tour in solution from list of cities in original_list, starting at start_index.
int tsp::nearest_neighbor_basic(int start_index)
{
    int cities_added = 0;
    int closest = 9999999;
    int total_dist = 0;
    int current_dist = 0;
    int closest_index = 0;
    int current_num = original_list.size();
    deque <city*> temp;
    copy_city_deque(original_list, temp);       //save original list

    solution.clear();
    solution.push_back(original_list[start_index]);     //move first city to solution
    original_list.erase(original_list.begin() + start_index);       //erase from original_list
    --current_num;        //cities remaining in original_list
    ++cities_added;       //number of cities in solution so far
    while(current_num != 0)             //loop until no cities remaining in original_list
    {
        closest = 9999999;  //reset closest to a large number so that comparison will work
        for (int i = 0; i < current_num; ++i)
        {
            current_dist = original_list[i]->dist(solution[cities_added-1]);
            if (current_dist < closest)
            {
                closest_index = i;
                closest = current_dist;
            }
        }

        total_dist += closest;
        solution.push_back(original_list[closest_index]);
        original_list.erase(original_list.begin() + closest_index);

        --current_num;
        ++cities_added;
        }

    copy_city_deque(temp, original_list);        //restore original list
    return total_dist + solution[0]->dist(solution[cities_added-1]);
}

//2-opt neighbor solution check
int tsp::two_change()
{
    deque <city*> new_path;
    int min_distance = get_solution_distance();
    bool start_over = false;

    signal(SIGTERM, end_opt);   //signal handler, ends optimization if it receives SIGTERM
    while(!done)
    {
        start_over = false;
        for (int i = 1; i < (int)original_list.size() && !start_over; ++i)
        {
            for (int j = i+1; j < (int)original_list.size()-1 && !start_over; ++j)
            {
                //only check moves that will reduce distance
                if (solution[i-1]->dist(solution[j]) + solution[i]->dist(solution[j+1]) < solution[i-1]->dist(solution[i]) + solution[j]->dist(solution[j+1]))
                {
                    swap_two(i, j);
                    min_distance = get_solution_distance();
                    start_over = true;
                }

                else
                    start_over = false;
            }
         }

         if (!start_over)
            break;
    }
    return min_distance;
}

//2-opt with neighbor list (NOT WORKING)
int tsp::two_opt()
{
    deque <city*> new_path;
    int min_distance = get_solution_distance();
    int k = 0;

        for (int i = 1; i < (int)original_list.size() ; ++i)
        {
            k = 1;
            fix_positions();
            while (k <= 5 && solution[i]->dist(solution[i]->get_neighbor(k)) < solution[i-1]->dist(solution[i]))
            {
                swap_two(i, solution[i]->get_neighbor_pos(k));
                min_distance = get_solution_distance();
                ++k;
                cout << min_distance << endl;
            }
            fix_positions();
        }


    return min_distance;
}

//Reverses the order of the cities from i to k.
int tsp::swap_two(int i, int k)
{
    deque <city*> temp;
    int count = 0;

    //Reverse order
    for (int x = k; x >= i; --x)
    {
        temp.push_back(solution[x]);
    }

    for (int x = i; x <= k; ++x)
    {
        solution[x] = temp[count];
        ++count;
    }
    temp.clear();
    //fix_positions();          //used for neighbor list algorithm
    return 1;
}

//Calculate the total distance of the tour in solution
int tsp::get_solution_distance()
{
    int total_dist = 0;
    for (int i = 0; i < (int)original_list.size() - 1; ++i)
    {
        total_dist += solution[i]->dist(solution[i+1]);
    }

    total_dist += solution[0]->dist(solution[original_list.size()-1]);
    return total_dist;
}

//Displays all the neighbor lists (used for testing)
void tsp::display_neighbor_lists()
{
    for (int i = 0; i < (int)original_list.size(); ++i)
     {
        cout << "LIST " << i << endl;
        solution[i]->display_neighbor_list();
    }

    return;
}

//Updates the stored positions of the cities to their current value
void tsp::fix_positions()
{
    for (int i = 0; i < (int)original_list.size(); ++i)
        solution[i]->set_pos(i);
}

//Copies a deque of city pointers from source to dest
void copy_city_deque(deque <city*> & source, deque <city*> & dest)
{
    int length = source.size();
    dest.clear();
    for (int i = 0; i < length; ++i)
        dest.push_back(new city(*source[i]));
}

//Catches the sigterm signal and updates global variable done to end optimization loops
void end_opt(int signum)
{
    cout << "\nOut of time\n";
    done = 1;
}
